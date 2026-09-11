import logging
import re
from collections import OrderedDict
from flask import Flask, jsonify, request
import sharepoint
import task_client
import request_client
import documents
import expense_client
import vector_memory
from azure_client import (
    ask_azure,
    ContentFilterError,
    generate_document_content,
    summarize_document,
    extract_task_details,
    extract_expense_details,
    extract_budget_details,
)
from commands import handle_command
from error_handler import handle_error
from vector_memory import (
    get_history,
    add_user_message,
    add_assistant_message,
)
from response_timer import (
    start_timer,
    stop_timer,
)
from speech_client import (
    transcribe_audio,
)
from calendar_client import (
    CALENDAR_USER_EMAIL,
    extract_calendar_details,
    add_calendar_event,
    get_calendar_event,
    update_calendar_event,
    parse_calendar_datetime,
    get_calendar_events,
    delete_calendar_event,
    check_for_conflict,
)
from typing_indicator import (
    send_typing_indicator,
)
from whatsapp import (
    VERIFY_TOKEN,
    send_message,
    send_document,
    download_document,
    download_audio,
)

from reminders import (
    handle_reminder_action,
    start_reminder_scheduler,
)
app = Flask(__name__)

logging.basicConfig(
    level=logging.INFO,
    format=(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(message)s"
    ),
)

log = logging.getLogger(
    __name__
)

# WhatsApp keeps retrying a webhook until it gets a 200 back, so the same
processed_messages = OrderedDict()
MAX_REMEMBERED_MESSAGES = 2000

def remember_message(message_id):
    processed_messages[message_id] = None

    while len(processed_messages) > MAX_REMEMBERED_MESSAGES:
        processed_messages.popitem(last=False)

# sender -> the slot we offered them, waiting on a "book it".
pending_conflicts = {}
CONFIRM_PHRASES = ["book it", "confirm", "yes", "sure", "okay", "ok"]
CANCEL_PHRASES = ["cancel", "never mind", "nevermind", "no"]

# sender -> the last task they touched, so "change the time of this task"
last_task_item_ids = {}

def clean_value(data, key, default=""):
    return str(data.get(key, default)).strip()

# Which day a message is about, defaulting to the whole upcoming window.
def day_in(words):
    if "tomorrow" in words:
        return "tomorrow"
    if "today" in words:
        return "today"
    return "upcoming"

# The number itself, however it is dressed up: "15", "ID 15", "#15", "no. 15".
TASK_NUMBER = r"(?:id|no|number)?\s*[#:]?\s*(\d+)"

# People write "delete task 15", but also "delete task ID 15" and
def parse_task_id(text_lower, command):
    match = re.fullmatch(
        TASK_NUMBER, text_lower.replace(command, "", 1).strip()
    )
    return match.group(1) if match else ""

# One request, several phrasings: complete task 2 | mark task 2 complete | change task 2 status to completed
COMPLETE_PATTERNS = [
    rf"^complete\s+task\s+{TASK_NUMBER}(?:\s+(?:please|now))?$",
    rf"^mark\s+task\s+{TASK_NUMBER}\s+(?:as\s+)?complete(?:d)?$",
    rf"^(?:change|set)\s+task\s+{TASK_NUMBER}"
    rf"\s+(?:status\s+)?(?:to\s+|as\s+)?complete(?:d)?$",
]

def parse_complete_task_id(text):
    normalized = " ".join(str(text).lower().split())

    for pattern in COMPLETE_PATTERNS:
        match = re.fullmatch(pattern, normalized)
        if match:
            return match.group(1)

    return ""

# Looser than the two above on purpose: the number can sit anywhere in the
def parse_reschedule_task_id(text):
    normalized = " ".join(
        str(text).lower().split()
    )
    match = re.search(rf"\btask\s+{TASK_NUMBER}\b", normalized)
    return match.group(1) if match else ""

# Keep the real SharePoint ID. The display number shifts whenever another
def remember_task(sender, task_result):
    item_id = str(
        (task_result or {}).get("id", "") or ""
    ).strip()

    if item_id:
        last_task_item_ids[sender] = item_id

# An explicit "task 3" always wins. Otherwise assume they mean the one they
def resolve_reschedule_task_id(sender, text):
    explicit_id = parse_reschedule_task_id(text)
    if explicit_id:
        return explicit_id

    item_id = last_task_item_ids.get(sender, "")
    if not item_id:
        return ""

    return task_client.get_display_id_for_item(
        item_id,
        sender,
    )

# Outlook first, then the task with the EventID already in the same POST.
def create_event_and_task(
    sender, task, title, date, start_time, end_time, priority, on_failure
):
    calendar_result = add_calendar_event(
        CALENDAR_USER_EMAIL, title, date, start_time, end_time
    )
    event_id = clean_value(calendar_result, "id")
    try:
        task_result = task_client.add_task(
            task, date, priority, sender, event_id=event_id
        )
    except Exception:
        log.exception(
            "Task creation failed after calendar event %s was created.",
            event_id,
        )
        if event_id:
            try:
                delete_calendar_event(CALENDAR_USER_EMAIL, event_id)
            except Exception:
                log.exception(
                    "Rollback also failed for calendar event %s", event_id
                )
        return on_failure

    remember_task(sender, task_result)
    display_id = task_client.get_display_id_for_item(
        task_result.get("id", ""), sender
    )

    reply = (
        "Task and calendar event created successfully.\n"
        f"Task: {task}\n"
        f"Date: {date}\n"
        f"Time: {start_time} - {end_time}\n"
        f"Priority: {priority}"
    )

    if display_id:
        reply += f"\nTask ID: {display_id}"

    return reply


def create_task_and_calendar(sender, text):
    task_details = extract_task_details(text)
    calendar_details = extract_calendar_details(text)

    task = clean_value(task_details, "task")
    priority = clean_value(task_details, "priority", "Medium")
    date = clean_value(calendar_details, "date") or clean_value(
        task_details, "due"
    )
    start_time = clean_value(calendar_details, "start_time")
    end_time = clean_value(calendar_details, "end_time")
    title = clean_value(calendar_details, "title", task) or task

    if not task:
        return "I could not determine the task name."
    if not date:
        return "I could not determine the task date."

    # No clock time means this is a to-do, not a meeting, so it goes to
    if not start_time or not end_time:
        task_result = task_client.add_task(task, date, priority, sender)
        remember_task(sender, task_result)
        display_id = task_client.get_display_id_for_item(
            task_result.get("id", ""), sender
        )

        reply = (
            "Task created successfully.\n"
            "(No specific time was given, so it was not added "
            "to your calendar.)\n"
            f"Task: {task}\n"
            f"Date: {date}\n"
            f"Priority: {priority}"
        )

        if display_id:
            reply += f"\nTask ID: {display_id}"

        return reply

    if not CALENDAR_USER_EMAIL:
        return (
            "I could not create the timed task because the calendar "
            "email is not configured. No task or event was created."
        )

    # Check for a clash before creating either record. Creating first and cleaning up afterwards leaves debris whenever the cleanup fails too.
    conflict = check_for_conflict(
        CALENDAR_USER_EMAIL, date, start_time, end_time
    )

    if conflict["conflict"]:
        reply = f"That time conflicts with \"{conflict['title']}\"."
        if conflict["suggested_start"]:
            reply += (
                f"\nThe next free slot that day is "
                f"{conflict['suggested_start']} - "
                f"{conflict['suggested_end']}.\n"
                "Reply \"book it\" if you want me to use that time instead."
            )
            pending_conflicts[sender] = {
                "mode": "create",
                "task": task,
                "title": title,
                "date": date,
                "start_time": conflict["suggested_start"],
                "end_time": conflict["suggested_end"],
                "priority": priority,
            }
        else:
            reply += "\nThere is no other free slot later that day."

        return reply

    return create_event_and_task(
        sender,
        task,
        title,
        date,
        start_time,
        end_time,
        priority,
        "I could not create the linked task. The calendar event was "
        "rolled back when possible. Make sure the WhatsApp Tasks list "
        "has an EventID column (Single line of text).",
    )

# Put an event back where it was, for when SharePoint fails after Outlook has already move
def _restore_calendar_event(user_email, event_id, event):
    if not event_id or not event:
        return

    old_start = parse_calendar_datetime(
        event.get("start", {}).get("dateTime", "")
    )
    old_end = parse_calendar_datetime(event.get("end", {}).get("dateTime", ""))

    if not old_start or not old_end:
        return

    update_calendar_event(
        user_email,
        event_id,
        event.get("subject", "Task"),
        old_start.strftime("%Y-%m-%d"),
        old_start.strftime("%H:%M"),
        old_end.strftime("%H:%M"),
    )

# Move a task in both systems, or in neither.
def apply_task_reschedule(sender, display_id, date, start_time, end_time):
    info, error = task_client.get_task_info(display_id, sender)

    if error:
        return error

    title = info["title"]
    old_event_id = info.get("event_id", "")
    old_event = None
    event_id = old_event_id
    created_new_event = False

    # Reuse the event that is already linked, so the link survives the move.
    if old_event_id:
        old_event = get_calendar_event(CALENDAR_USER_EMAIL, old_event_id)

    if old_event:
        updated = update_calendar_event(
            CALENDAR_USER_EMAIL,
            old_event_id,
            title,
            date,
            start_time,
            end_time,
        )

        # The saved EventID points at something that is gone (deleted in Outlook directly, usually). Treat the task as unlinked and make a fresh event.
        if updated is None:
            old_event = None
            event_id = ""

    if not old_event:
        calendar_result = add_calendar_event(
            CALENDAR_USER_EMAIL, title, date, start_time, end_time
        )
        event_id = clean_value(calendar_result, "id")
        created_new_event = True

    # Outlook has changed by this point, so anything that fails from here has to be undone
    try:
        update_result, update_error = task_client.update_task_schedule(
            display_id, sender, date, event_id=event_id
        )

        if update_error:
            raise RuntimeError(update_error)

    except Exception as error:
        log.exception(
            "Could not synchronize SharePoint after rescheduling task %s.",
            display_id,
        )

        # Roll back whichever half we did, so the two do not drift apart.
        try:
            if created_new_event and event_id:
                delete_calendar_event(CALENDAR_USER_EMAIL, event_id)
            elif old_event and old_event_id:
                _restore_calendar_event(
                    CALENDAR_USER_EMAIL, old_event_id, old_event
                )
        except Exception:
            log.exception("Calendar rollback failed for task %s.", display_id)

        return (
            "I could not synchronize the task and Outlook calendar, so I "
            "rolled the calendar change back when possible. "
            f"Details: {error}"
        )

    if update_result:
        last_task_item_ids[sender] = update_result["item_id"]

    return (
        f'Task {display_id} "{title}" was rescheduled successfully.\n'
        f"Date: {date}\n"
        f"Time: {start_time} - {end_time}\n"
        "Outlook Calendar and SharePoint are synchronized."
    )

# "change task 2 to 3 PM" - work out which task and which time, then check the calendar before touching anything.
def reschedule_task_and_calendar(sender, text):
    display_id = resolve_reschedule_task_id(sender, text)

    if not display_id:
        return (
            "Please include the task ID, for example: "
            '"change task 2 time to 1 PM to 2 PM".'
        )

    info, error = task_client.get_task_info(display_id, sender)

    if error:
        return error

    details = extract_calendar_details(text)
    date = clean_value(details, "date") or info.get("due", "")
    start_time = clean_value(details, "start_time")
    end_time = clean_value(details, "end_time")

    if not date:
        return "I could not determine the date for this task."

    if not start_time or not end_time:
        return (
            "Please include the new time, for example: "
            '"change task 2 time to 1 PM to 2 PM".'
        )

    # Fresh lookup on every command. Passing the task's own event means it does not count as a clash with itself.
    conflict = check_for_conflict(
        CALENDAR_USER_EMAIL,
        date,
        start_time,
        end_time,
        ignore_event_id=info.get("event_id", ""),
    )

    if conflict["conflict"]:
        reply = (
            f'Task {display_id} was not changed. '
            f'That time conflicts with "{conflict["title"]}".'
        )

        if conflict.get("suggested_start"):
            pending_conflicts[sender] = {
                "mode": "reschedule",
                "display_id": display_id,
                "item_id": info["item_id"],
                "title": info["title"],
                "event_id": info.get("event_id", ""),
                "date": date,
                "start_time": conflict["suggested_start"],
                "end_time": conflict["suggested_end"],
            }
            reply += (
                f"\nThe next free slot is "
                f"{conflict['suggested_start']} - "
                f"{conflict['suggested_end']}.\n"
                'Reply "book it" to move the task to that slot.'
            )
        else:
            reply += "\nThere is no other free slot later that day."

        return reply

    return apply_task_reschedule(sender, display_id, date, start_time, end_time)

# They replied "book it". Minutes may have passed since the slot was offered, so ask Outlook again rather than trusting what we suggested earlier.
def book_pending_conflict(sender):
    pending = pending_conflicts.get(sender)

    if not pending:
        return "There is no pending booking to confirm."

    mode = pending.get("mode", "create")
    ignore_event_id = pending.get("event_id", "") if mode == "reschedule" else ""

    conflict = check_for_conflict(
        CALENDAR_USER_EMAIL,
        pending["date"],
        pending["start_time"],
        pending["end_time"],
        ignore_event_id=ignore_event_id,
    )

    if conflict["conflict"]:
        if conflict.get("suggested_start"):
            pending["start_time"] = conflict["suggested_start"]
            pending["end_time"] = conflict["suggested_end"]
            return (
                f'That suggested slot is now busy because of '
                f'"{conflict["title"]}".\n'
                f"The next free slot is "
                f"{conflict['suggested_start']} - "
                f"{conflict['suggested_end']}.\n"
                'Reply "book it" again if you want that new slot.'
            )

        pending_conflicts.pop(sender, None)
        return (
            "That suggested slot is no longer available, and I could not "
            "find another free slot later that day."
        )

    pending_conflicts.pop(sender, None)

    if mode == "reschedule":
        return apply_task_reschedule(
            sender,
            pending["display_id"],
            pending["date"],
            pending["start_time"],
            pending["end_time"],
        )

    return create_event_and_task(
        sender,
        pending["task"],
        pending["title"],
        pending["date"],
        pending["start_time"],
        pending["end_time"],
        pending.get("priority", "Medium"),
        "I could not create the linked task, so I rolled back the "
        "calendar booking when possible.",
    )

DOCUMENT_PHRASES = [
    "create a document", "create document",
    "generate a document", "generate document",
]

SHAREPOINT_KEYWORDS = [
    "company", "policy", "policies", "vacation", "remote", "working hours",
    "dress code", "department", "manager", "budget", "report", "employee",
]

# Rough keyword check. It saves reading the whole SharePoint site into the prompt for messages that clearly have nothing to do with the company.
def needs_sharepoint(text):
    text = text.lower()

    return any(keyword in text for keyword in SHAREPOINT_KEYWORDS)

# Everything a user types ends up here. The order of the branches below is load-bearing - see the note above the calendar ones.
def process_text(sender, text):
    text_lower = " ".join(text.lower().split())
    # A set of whole words, so a check for "no" cannot fire on "nothing" and "task" cannot fire on "tasks" unless we ask for both.
    words = set(re.findall(r"[a-z]+", text_lower))
    send_text_reply = True

    # Catch completion requests before they fall through to the AI. A number is still required further down before SharePoint is touched.
    complete_task_requested = (
        "task" in words
        and words & {"complete", "completed"}
        and words & {"complete", "mark", "change", "set"}
    )

    # "change task 2 to 3 PM" never says the word time or date, so a bare clocktime has to count as a reschedule on its own.
    has_clock_time = bool(
        re.search(
            r"\b\d{1,2}:\d{2}\b|"
            r"\b\d{1,2}(?::\d{2})?\s*(?:a\.?m\.?|p\.?m\.?)\b",
            text_lower,
        )
    )

    # "change task 2 status to complete" also contains "change", which is why the completion case above is excluded here.
    reschedule_task_requested = (
        "task" in words
        and not complete_task_requested
        and (
            "reschedule" in words
            or "move" in words
            or (
                bool(words & {"change", "update", "set"})
                and (
                    "time" in words
                    or "date" in words
                    or has_clock_time
                )
            )
        )
    )

    # "schedule" is the awkward one: a verb in "schedule a meeting", a noun in "show my schedule". These words say they are reading, not creating.
    calendar_read_words = {
        "show", "list", "what", "all", "see", "view", "upcoming", "have",
    }
    creation_entities = {
        "task", "tasks", "calendar", "event", "events", "meeting",
        "appointment",
    }
    explicit_create_words = {"add", "create", "book"}

    calendar_read_requested = (
        "calendar" in words and words & calendar_read_words
    )

    # So creating needs something to create, plus either a plain verb or "schedule" with none of those reading words nearby.
    create_task_or_calendar_requested = words & creation_entities and (
        words & explicit_create_words
        or ("schedule" in words and not words & calendar_read_words)
    )

    # Confirming the slot we offered earlier.
    if sender in pending_conflicts and any(
        phrase in text_lower for phrase in CONFIRM_PHRASES
    ):
        reply = book_pending_conflict(sender)

    # Backing out of it.
    elif sender in pending_conflicts and any(
        phrase in text_lower for phrase in CANCEL_PHRASES
    ):
        pending_conflicts.pop(sender, None)
        reply = "Okay, I will not book that."

    # A plain greeting is as good a moment as any to show them the day ahead.
    elif text_lower in ("hi", "hello", "hey"):
        reply = (
            "Hello! I can answer questions using "
        "information from the SharePoint site, "
        "create and upload documents, upload "
        "documents you send me, create requests, "
        "and create tasks that are added to your "
        "Outlook calendar. I can also track your "
        "expenses - just tell me what you spent, "
        "and ask me anytime how much you've "
        "spent so far.\n\n"
        "You can also give feedback here:\n"
        "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=90ZqoA_Yv0KHZw9pHZRS6hU5uAcUAndDiA6G7lmnLr5UQkJPSEFIS1UwN0pBTk5WREFFVDRLTVM1Ui4u"
    )

        try:
            reply += (
                "\n\nHere is what you have today:\n"
                f"{get_calendar_events(CALENDAR_USER_EMAIL, 'today')}\n\n"
                f"{task_client.get_tasks_due_on(sender, 'today')}"
            )
        except Exception:
            log.exception(
                "Could not load today's schedule/tasks "
                "for the hi/hello greeting."
            )

    # A request typed as name/department/request, one field per line.
    elif text_lower.startswith(
        "add request"
    ):
        fields = {
            "name": "",
            "department": "",
            "request": "",
        }

        for line in text.splitlines():
            if ":" not in line:
                continue
            key, value = line.split(
                ":",
                1,
            )
            key = key.strip().lower()
            if key in fields:
                fields[key] = value.strip()

        request_client.add_request(
            fields["name"],
            fields["department"],
            fields["request"],
            sender,
        )
        reply = "Your request was added to SharePoint."

    elif text_lower == "check my request":
        reply = request_client.get_request_status(sender)

    # "show my tasks", but also "what are all my tasks" and "list my tasks". Calendar and today/tomorrow are ruled out here so the narrower branches further down still get their turn at those.
    elif (
        not words & {"calendar", "today", "tomorrow"}
        and words & {"task", "tasks"}
        and words & {"show", "list", "what", "all", "see", "view"}
    ):
        reply = task_client.get_tasks(sender)

    # Mark it done. This never deletes anything.
    elif complete_task_requested:
        task_id = parse_complete_task_id(text)

        if not task_id:
            reply = (
                "Please tell me which task to complete.\n"
                "For example:\n"
                "complete task 2\n"
                "mark task 2 complete\n"
                "change task 2 status to complete"
            )
        else:
            reply = task_client.complete_task(task_id, sender)

    # Move an existing task. Outlook is read live before either side changes.
    elif reschedule_task_requested:
        reply = reschedule_task_and_calendar(sender, text)

    # Deletion stays an exact command. No amount of loose phrasing should be able to trigger it by accident.
    elif text_lower.startswith("delete task"):
        task_id = parse_task_id(text_lower, "delete task")

        if not task_id:
            reply = "Please include the task ID, for example: delete task 15"
        else:
            reply = task_client.delete_task(task_id, sender)

    # Calendar reads have to sit above the creation branch. "show my upcoming calendar schedule" contains the word schedule and would otherwise be read as a request to book something.
    elif calendar_read_requested:
        reply = get_calendar_events(CALENDAR_USER_EMAIL, day_in(words))

    # Creates the SharePoint task and the Outlook event as one unit.
    elif create_task_or_calendar_requested:
        reply = create_task_and_calendar(sender, text)

    # Tasks for a given day. If they actually said "calendar", the branches above have already answered.
    elif (
        words & {"today", "tomorrow"}
        and "calendar" not in words
        and words & {"task", "tasks"}
    ):
        reply = task_client.get_tasks_due_on(sender, day_in(words))

    # "what do I have today?" is vague enough that they probably want both lists.
    elif words & {"today", "tomorrow"} and "have" in words:
        day = day_in(words)
        reply = (
            get_calendar_events(CALENDAR_USER_EMAIL, day)
            + "\n\n"
            + task_client.get_tasks_due_on(sender, day)
        )

    # Event IDs are long and case-sensitive, and they come from an earlier reply, so read them out of the original text rather than the lowercased copy.
    elif text_lower.startswith("delete event "):
        event_id = " ".join(text.split())[len("delete event "):].strip()
        if not event_id:
            reply = "Please include the event ID."
        else:
            reply = delete_calendar_event(CALENDAR_USER_EMAIL, event_id)

    # Generate a Word document and send it back as a file.
    elif any(phrase in text_lower for phrase in DOCUMENT_PHRASES):
        # Never write another employee's salary into a document.
        if "salary" in text_lower and "my salary" not in text_lower:
            reply = (
                "I cannot create a document containing another "
                "employee's salary information."
            )
        else:
            if "my salary" in text_lower:
                document_content = sharepoint.get_my_salary(sender)
            else:
                document_content = generate_document_content(
                    text,
                    documents.read_sharepoint() if needs_sharepoint(text) else "",
                )

            filename = documents.create_document(text, document_content)
            documents.upload_document(filename)
            send_document(sender, filename)

            reply = f"Document created: {filename}"
            send_text_reply = False

    # Salary is only ever answered for the number that asked.
    elif "salary" in text_lower:
        if "my salary" in text_lower:
            reply = sharepoint.get_my_salary(sender)
        else:
            reply = (
                "I can only provide salary information for the employee "
                "associated with your WhatsApp number."
            )

    # Reading expenses has to be checked before logging them, or "how much have I spent" trips the branch below on the word spent.
    elif "how much" in text_lower or (
        words & {"expense", "expenses"}
        and words & {"show", "what", "all", "view", "see"}
    ):
        reply = expense_client.get_expenses_this_month(sender)

    # Log a new expense.
    elif words & {"spent", "spend", "bought", "paid", "purchased", "purchase"}:
        expense_details = extract_expense_details(text)

        description = clean_value(expense_details, "description")
        amount = clean_value(expense_details, "amount")
        category = clean_value(expense_details, "category", "Other")

        if not description:
            reply = "I could not determine what the expense was for."
        elif not amount:
            reply = "I could not determine the expense amount."
        else:
            expense_client.add_expense(description, amount, category, sender)
            reply = (
                "Expense logged successfully.\n"
                f"Description: {description}\n"
                f"Amount: ${amount}\n"
                f"Category: {category}"
            )

            budgets = expense_client.get_budgets(sender)

            # Only worth the extra call when a budget exists for this category.
            if category in budgets:
                _, by_category = expense_client.get_month_totals(sender)
                spent = by_category.get(category, 0)
                limit = budgets[category]["amount"]

                if spent > limit:
                    reply += (
                        f"\n\n⚠️ You have gone over your {category} budget "
                        f"of ${limit:.2f} this month (spent ${spent:.2f})."
                    )

    # Set a monthly limit.
    elif "budget" in words and words & {"set", "update", "change"}:
        budget_details = extract_budget_details(text)

        # One message can set several budgets at once, and the model returns a list when it does.
        if not isinstance(budget_details, list):
            budget_details = [budget_details]

        confirmations = []

        for entry in budget_details:
            budget_category = clean_value(
                entry,
                "category",
            )
            budget_amount = clean_value(
                entry,
                "amount",
            )

            if not budget_category or not budget_amount:
                continue

            expense_client.set_budget(budget_category, budget_amount, sender)
            confirmations.append(
                f"- {budget_category}: ${budget_amount} per month"
            )

        if not confirmations:
            reply = "I could not determine the budget category or amount."
        else:
            reply = "Budgets set:\n" + "\n".join(confirmations)

    # Show the limits.
    elif "budget" in words and words & {"show", "what", "view", "see"}:
        budgets = expense_client.get_budgets(sender)

        if not budgets:
            reply = "You have not set any budgets yet."
        else:
            reply = "\n".join(
                ["Your monthly budgets:"]
                + [
                    f"- {category}: ${info['amount']:.2f}"
                    for category, info in budgets.items()
                ]
            )

    # Nothing matched, so try the explicit commands like /clear, and fall back to a plain AI answer if that is not one either.
    else:
        command_reply = handle_command(sender, text)

        if command_reply is not None:
            reply = command_reply
        else:
            reply = ask_azure(
                text,
                get_history(sender)[-6:],
                documents.read_sharepoint() if needs_sharepoint(text) else "",
                vector_memory.recall_context(sender, text),
            )

    add_user_message(sender, text)
    add_assistant_message(sender, reply)

    if send_text_reply:
        send_message(
            sender,
            reply,
        )

    return reply

# Health check, so the host can tell the app is alive.
@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "running",
        "message": "WhatsApp AI assistant is running",
    }), 200

# Every path answers WhatsApp the same way: 200, so it stops retrying.
def ack():
    return jsonify({"status": "received"}), 200

# Everything from WhatsApp arrives here.
@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    # Meta calls this once with a token to verify the endpoint.
    if request.method == "GET":
        if request.args.get("hub.verify_token") == VERIFY_TOKEN:
            return request.args.get("hub.challenge"), 200
        return "Forbidden", 403

    payload = request.get_json(silent=True) or {}
    sender = None

    try:
        value = (
            payload
            .get("entry", [{}])[0]
            .get("changes", [{}])[0]
            .get("value", {})
        )

        # Delivery receipts and read receipts come through this webhook too.
        if "messages" not in value:
            return ack()

        message = value["messages"][0]
        sender = message.get("from")
        message_id = message.get("id")
        message_type = message.get("type")

        # Already answered this one on an earlier retry.
        if message_id and message_id in processed_messages:
            return ack()

        if message_id:
            remember_message(message_id)

        if message_type == "document":
            document = message.get("document", {})
            media_id = document.get("id")
            filename = document.get("filename", "whatsapp_document")
            mime_type = document.get(
                "mime_type", "application/octet-stream"
            )

            if not media_id:
                raise ValueError("WhatsApp media ID is missing.")

            if message_id:
                send_typing_indicator(message_id)

            start_time = start_timer()

            #download document
            file_content = download_document(media_id)

            documents.upload_whatsapp_document(
                filename, file_content, mime_type
            )
            document_text = documents.extract_document_text(
                filename, file_content
            )

            if not document_text.strip():
                raise ValueError(
                    "No readable text was found in the uploaded document."
                )

            log.info(
                "Extracted document text length: %d characters",
                len(document_text),
            )

            send_message(
                sender,
                f"{filename} was uploaded successfully to SharePoint.\n\n"
                f"Summary:\n{summarize_document(document_text)}",
            )
            log.info(
                "Document processing time: %.2f seconds",
                stop_timer(start_time),
            )

            return ack()

        if message_type == "audio":
            audio = message.get("audio", {})
            media_id = audio.get("id")

            if not media_id:
                raise ValueError("WhatsApp audio ID is missing.")

            if message_id:
                send_typing_indicator(message_id)

            start_time = start_timer()
            audio_bytes, detected_mime = download_audio(media_id)
            transcript = transcribe_audio(
                audio_bytes,
                detected_mime or audio.get("mime_type", "audio/ogg"),
            )

            log.info("Voice transcript from %s: %s", sender, transcript)
            process_text(sender, transcript)
            log.info(
                "Voice response time: %.2f seconds", stop_timer(start_time)
            )

            return ack()

        # Snooze / Dismiss taps on a reminder. T
        if message_type == "interactive":
            button_id = (
                message
                .get("interactive", {})
                .get("button_reply", {})
                .get("id", "")
            )

            if ":" in button_id:
                action, event_id = button_id.split(":", 1)
                reminder_reply = handle_reminder_action(
                    action, event_id, sender
                )

                if reminder_reply:
                    send_message(sender, reminder_reply)

            return ack()

        if message_type != "text":
            return ack()

        text = message.get("text", {}).get("body", "").strip()

        if not text:
            return ack()

        log.info("Message from %s: %s", sender, text)

        if message_id:
            send_typing_indicator(message_id)

        start_time = start_timer()
        process_text(sender, text)
        log.info("Response time: %.2f seconds", stop_timer(start_time))

    except Exception as error:
        if not isinstance(error, ContentFilterError):
            log.exception("Webhook error.")

        if sender:
            try:
                send_message(sender, handle_error(error))
            except Exception as send_error:
                log.error(
                    "Could not send the error message to %s: %s",
                    sender,
                    send_error,
                )

    return ack()

if __name__ == "__main__":
    log.info("Starting WhatsApp AI assistant...")
    start_reminder_scheduler()
    app.run(
        host="0.0.0.0",
        port=3000,
        debug=True,
        use_reloader=False,
    )
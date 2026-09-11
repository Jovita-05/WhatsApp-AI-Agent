from datetime import datetime, timedelta
import logging
import sharepoint
from sharepoint import json_headers, list_items_url, normalize_number
from calendar_client import (
    TIMEZONE,
    CALENDAR_USER_EMAIL,
    delete_calendar_event,
    find_event_id,
    to_local_date,
)
#preparation code
log = logging.getLogger(__name__)

TASK_LIST = "WhatsApp Tasks"
INVALID_ID = 'That task ID is not valid. Send "my tasks" to see the IDs.'
NOT_FOUND = 'Task was not found. Send "my tasks" to see the current IDs.'

def get_task_context():
    return sharepoint.get_list_context(TASK_LIST)

def _get_all_list_items(headers, site_id, list_id):
    url = list_items_url(site_id, list_id) + "?expand=fields"

    response = sharepoint.session.get(
        url,
        headers=headers,
        timeout=30
    )
    response.raise_for_status()
    data = response.json()

    return data.get("value", [])

#all tasks
def _get_user_items(whatsapp_number):
    headers, site_id, list_id = get_task_context()
    user_number = normalize_number(whatsapp_number)
    all_items = _get_all_list_items(headers, site_id, list_id)
    user_items = []

    for item in all_items:
        fields = item.get("fields", {})

        saved_number = fields.get("WhatsAppNumber", "")
        saved_number = normalize_number(saved_number)

        if saved_number == user_number:
            user_items.append(item)
            #important for sorting tasks order
    user_items.sort(key=lambda item: int(item.get("id", 0)))
    return headers, site_id, list_id, user_items

#takes all tasks and find a specific one
def get_user_task(display_id, whatsapp_number):
    task_number_text = str(display_id).strip()
    if not task_number_text.isdigit():
        return None, INVALID_ID

    task_number = int(task_number_text)

    if task_number < 1:
        return None, INVALID_ID

    headers, site_id, list_id, user_items = _get_user_items(whatsapp_number)

    if task_number > len(user_items):
        return None, NOT_FOUND
    item = user_items[task_number - 1]
    item_id = str(item.get("id", ""))
    fields = item.get("fields", {})

    context = (
        headers,
        site_id,
        list_id,
        item_id,
        fields,
        item,
        task_number,
    )

    return context, None

#convert
def get_display_id_for_item(item_id, whatsapp_number):
    _, _, _, user_items = _get_user_items(whatsapp_number)
    display_id = 1

    for item in user_items:
         if str(item.get("id", "")) == str(item_id):
            return str(display_id)

         display_id += 1

    return ""

#add a task
def add_task(task, due_date, priority, whatsapp_number, event_id=""):
    headers, site_id, list_id = get_task_context()

    fields = {
        "Title": task,
        "DueDate": due_date,
        "Priority": priority,
        "Status": "Pending",
        "WhatsAppNumber": str(whatsapp_number),
    }

    if event_id:
        fields["EventID"] = event_id

    url = list_items_url(site_id, list_id)

    response = sharepoint.session.post(
        url,
        headers=json_headers(headers),
        json={"fields": fields},
        timeout=60,
    )

    response.raise_for_status()
    return response.json()

#for updated stuff
def _patch_fields(headers, site_id, list_id, item_id, updates):
    fields_url = list_items_url(site_id, list_id, item_id) + "/fields"

    response = sharepoint.session.patch(
        fields_url,
        headers=json_headers(headers),
        json=updates,
        timeout=30,
    )

    response.raise_for_status()
    verify_response = sharepoint.session.get(
        fields_url,
        headers=headers,
        timeout=30,
    )

    verify_response.raise_for_status()
    return verify_response.json()

#collect imp info
def get_task_info(display_id, whatsapp_number):
    context, error = get_user_task(display_id, whatsapp_number)

    if error:
        return None, error

    (
        headers,
        site_id,
        list_id,
        item_id,
        fields,
        item,
        current_display_id,
    ) = context

    task_info = {
        "display_id": str(current_display_id),
        "item_id": str(item_id),
        "title": str(fields.get("Title", "Task") or "Task"),
        "due": to_local_date(fields.get("DueDate", "")),
        "priority": str(fields.get("Priority", "Medium") or "Medium"),
        "status": str(fields.get("Status", "Pending") or "Pending"),
        "event_id": str(fields.get("EventID", "") or "").strip(),
        "created_at": str(item.get("createdDateTime", "") or ""),
    }

    return task_info, None

#update and call patch 
def update_task_schedule(
    display_id,
    whatsapp_number,
    due_date,
    event_id=None,
):
    context, error = get_user_task(display_id, whatsapp_number)

    if error:
        return None, error

    (
        headers,
        site_id,
        list_id,
        item_id,
        fields,
        item,
        current_display_id,
    ) = context

    updates = {
        "DueDate": due_date,
    }

    if event_id is not None:
        updates["EventID"] = event_id

    saved_fields = _patch_fields(
        headers,
        site_id,
        list_id,
        item_id,
        updates,
    )

    saved_due_date = to_local_date(saved_fields.get("DueDate", ""))
    saved_event_id = str(saved_fields.get("EventID", "") or "").strip()

    if saved_due_date != due_date:
        return None, (
            "The calendar changed, but I could not verify the new "
            "task date in SharePoint."
        )

    if event_id is not None:
        if saved_event_id != str(event_id):
            return None, (
                "The calendar changed, but I could not verify the "
                "Outlook event link in SharePoint."
            )

    result = {
        "display_id": str(current_display_id),
        "item_id": str(item_id),
        "due": saved_due_date,
        "event_id": saved_event_id,
    }

    return result, None

#
def get_tasks(whatsapp_number):
    headers, site_id, list_id, user_items = _get_user_items(whatsapp_number)

    if not user_items:
        return "You do not have any tasks."

    lines = ["Your tasks:"]

    display_id = 1

    for item in user_items:
        fields = item.get("fields", {})

        title = fields.get("Title", "")
        due_date = to_local_date(fields.get("DueDate", ""))
        priority = fields.get("Priority", "")
        status = fields.get("Status", "Pending") or "Pending"

        task_text = (
            f"\n{display_id}. {title}\n"
            f"Due: {due_date}\n"
            f"Priority: {priority}\n"
            f"Status: {status}"
        )

        lines.append(task_text)

        display_id += 1

    return "\n".join(lines)

#
def get_tasks_due_on(whatsapp_number, day="today"):
    now = datetime.now(TIMEZONE)

    if day == "today":
        due_date = now.strftime("%Y-%m-%d")

    elif day == "tomorrow":
        tomorrow = now + timedelta(days=1)
        due_date = tomorrow.strftime("%Y-%m-%d")

    else:
        due_date = day

    headers, site_id, list_id, user_items = _get_user_items(whatsapp_number)

    matching_tasks = []

    for item in user_items:
        fields = item.get("fields", {})

        item_due_date = to_local_date(fields.get("DueDate", ""))

        if item_due_date == due_date:
            title = fields.get("Title", "")
            priority = fields.get("Priority", "")
            status = fields.get("Status", "Pending") or "Pending"

            task_text = (
                f"- {title} "
                f"(Priority: {priority}, Status: {status})"
            )

            matching_tasks.append(task_text)

    if not matching_tasks:
        return f"You have no tasks due {day}."

    heading = f"Tasks due {day}:"

    return "\n".join([heading] + matching_tasks)

#complete
def complete_task(display_id, whatsapp_number):
    context, error = get_user_task(display_id, whatsapp_number)

    if error:
        return error

    (
        headers,
        site_id,
        list_id,
        item_id,
        fields,
        item,
        current_display_id,
    ) = context

    title = str(fields.get("Title", "Task"))

    updates = {
        "Status": "Completed",
    }

    saved_fields = _patch_fields(
        headers,
        site_id,
        list_id,
        item_id,
        updates,
    )

    saved_status = str(saved_fields.get("Status", "")).strip()

    if saved_status.casefold() != "completed":
        log.error(
            "Task %s completion PATCH returned success but "
            "Status read back as %r",
            item_id,
            saved_status,
        )

        return (
            "I could not verify that the task was completed. "
            "Please check the SharePoint Status column configuration."
        )

    return (
        f'Task {current_display_id} "{title}" marked as completed.'
    )

#delete
def delete_task(display_id, whatsapp_number):
    context, error = get_user_task(display_id, whatsapp_number)

    if error:
        return error

    (
        headers,
        site_id,
        list_id,
        item_id,
        fields,
        item,
        current_display_id,
    ) = context

    title = str(fields.get("Title", "Task"))
    event_id = str(fields.get("EventID", "") or "").strip()

    if not event_id and CALENDAR_USER_EMAIL:
        try:
            event_id = find_event_id(
                CALENDAR_USER_EMAIL,
                title,
                to_local_date(fields.get("DueDate", "")),
                str(item.get("createdDateTime", "")),
            )

        except Exception:
            log.exception(
                "Could not look up the calendar event for task %s",
                item_id,
            )

            event_id = ""

    calendar_result = ""

    if event_id and CALENDAR_USER_EMAIL:
        try:
            calendar_result = delete_calendar_event(
                CALENDAR_USER_EMAIL,
                event_id,
            )

        except Exception:
            log.exception(
                "Could not delete calendar event %s for task %s",
                event_id,
                item_id,
            )

            return (
                "I could not remove the linked Outlook event, so I did not "
                "delete the task. Please try again."
            )

        valid_calendar_results = {
            "Calendar event deleted.",
            "Calendar event was not found.",
        }

        if calendar_result not in valid_calendar_results:
            return (
                "I could not confirm removal of the linked Outlook event, "
                "so I did not delete the task."
            )

    task_url = list_items_url(
        site_id,
        list_id,
        item_id,
    )

    response = sharepoint.session.delete(
        task_url,
        headers=headers,
        timeout=30,
    )

    response.raise_for_status()

    if calendar_result == "Calendar event deleted.":
        note = ", and its calendar event was removed too."

    elif calendar_result == "Calendar event was not found.":
        note = ". The linked calendar event was already gone."

    else:
        note = "."

    return (
        f'Task {current_display_id} "{title}" deleted{note} '
        "Task IDs have been refreshed."
    )
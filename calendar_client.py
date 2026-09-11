import os
import logging
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import sharepoint
from azure_client import get_json_response

#prep
CALENDAR_USER_EMAIL = os.getenv("CALENDAR_USER_EMAIL", "").strip()
CALENDAR_TIMEZONE = "Asia/Beirut"
TIMEZONE = ZoneInfo(CALENDAR_TIMEZONE)
UTC = ZoneInfo("UTC")
SLOT_STEP_MINUTES = 30
LATEST_SUGGESTION_HOUR = 20  # nobody wants a meeting suggested at 11 PM

log = logging.getLogger(__name__)

def parse_iso(value, default_tz):
    if not value:
        return None

    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None

    return parsed.replace(tzinfo=default_tz) if parsed.tzinfo is None else parsed

def parse_calendar_datetime(value):
    parsed = parse_iso(value, TIMEZONE)

    return parsed.astimezone(TIMEZONE) if parsed else None

def parse_graph_time(value):
    return parse_iso(value, UTC)

def to_local_date(value):
    if not value:
        return ""

    parsed = parse_iso(value, UTC)

    if parsed is None:
        return str(value)[:10]

    return parsed.astimezone(TIMEZONE).strftime("%Y-%m-%d")

#natural language 
def extract_calendar_details(text):
    now = datetime.now(TIMEZONE)

    result = get_json_response(
        "Extract calendar event information "
        "from the user's message. "
        f"Today's date is {now.strftime('%Y-%m-%d')} "
        f"({now.strftime('%A')}). Use this to resolve "
        "words like 'today' or 'tomorrow'. "
        "Return ONLY valid JSON with these keys: "
        "title, date, start_time, end_time. "
        "Use YYYY-MM-DD for date. "
        "Use HH:MM 24-hour format for times. "
        "If the message does not mention a specific "
        "start time, return an empty string for "
        "start_time and end_time - do not invent a time. "
        "If a start time is given but no end time, "
        "make the event one hour long. "
        "If that one-hour event crosses midnight, "
        "end_time may be 00:00.",
        text,
    )

    if isinstance(result, list):
        result = result[0] if result else {}
    if not isinstance(result, dict):
        raise ValueError(
            "Calendar details were not returned in the expected format."
        )
    return result

# The Prefer header is what makes Graph return times in Beirut instead of UTC.
def get_calendar_headers():
    headers, _ = sharepoint.get_graph_context()

    return {
        **headers,
        "Content-Type": "application/json",
        "Prefer": f'outlook.timezone="{CALENDAR_TIMEZONE}"',
    }

def events_url(user_email, event_id=""):
    url = f"{sharepoint.GRAPH_URL}/users/{user_email}/events"

    return f"{url}/{event_id}" if event_id else url

#create start / end time and understand 12am
def build_event_datetimes(date, start_time, end_time):
    start_dt = datetime.strptime(
        f"{date} {start_time}", "%Y-%m-%d %H:%M"
    ).replace(tzinfo=TIMEZONE)
    end_dt = datetime.strptime(
        f"{date} {end_time}", "%Y-%m-%d %H:%M"
    ).replace(tzinfo=TIMEZONE)

    if end_dt <= start_dt:
        end_dt += timedelta(days=1)

    return start_dt, end_dt

# The start/end block every event write sends to Graph.
def event_times(date, start_time, end_time):
    start_dt, end_dt = build_event_datetimes(date, start_time, end_time)

    return {
        "start": {
            "dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": CALENDAR_TIMEZONE,
        },
        "end": {
            "dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": CALENDAR_TIMEZONE,
        },
    }

# Create an Outlook calendar event
def add_calendar_event(user_email, title, date, start_time, end_time):
    if not user_email:
        raise ValueError("Calendar email is not configured.")

    response = sharepoint.session.post(
        events_url(user_email),
        headers=get_calendar_headers(),
        json={"subject": title, **event_times(date, start_time, end_time)},
        timeout=30,
    )
    response.raise_for_status()

    return response.json()

# retrieving a specific calendar event when I know its event ID
def get_calendar_event(user_email, event_id):
    if not user_email or not event_id:
        return None

    response = sharepoint.session.get(
        events_url(user_email, event_id),
        headers=get_calendar_headers(),
        timeout=30,
    )

    if response.status_code == 404:
        return None

    response.raise_for_status()

    return response.json()

# update
def update_calendar_event(
    user_email, event_id, title, date, start_time, end_time
):
    if not user_email:
        raise ValueError("Calendar email is not configured.")
    if not event_id:
        raise ValueError("Calendar event ID is missing.")

    response = sharepoint.session.patch(
        events_url(user_email, event_id),
        headers=get_calendar_headers(),
        json={"subject": title, **event_times(date, start_time, end_time)},
        timeout=30,
    )

    if response.status_code == 404:
        return None

    response.raise_for_status()

    if response.content:
        try:
            return response.json()
        except ValueError:
            pass

    return get_calendar_event(user_email, event_id)

# Delete an Outlook calendar event
def delete_calendar_event(user_email, event_id):
    if not user_email:
        return "Calendar email is not configured."

    response = sharepoint.session.delete(
        events_url(user_email, event_id),
        headers=get_calendar_headers(),
        timeout=30,
    )

    if response.status_code == 404:
        return "Calendar event was not found."

    response.raise_for_status()

    return "Calendar event deleted."

# Everything on the calendar between two moments, oldest first.
def get_calendar_view(user_email, start, end):
    response = sharepoint.session.get(
        f"{sharepoint.GRAPH_URL}/users/{user_email}/calendarView",
        headers=get_calendar_headers(),
        params={
            "startDateTime": start.isoformat(),
            "endDateTime": end.isoformat(),
            "$orderby": "start/dateTime",
        },
        timeout=30,
    )
    response.raise_for_status()

    return response.json().get("value", [])

#conflict checking
def times_overlap(start_a, end_a, start_b, end_b):
    return start_a < end_b and end_a > start_b

# Is the requested slot free? 
def check_for_conflict(
    user_email, date, start_time, end_time, ignore_event_id=""
):
    if not user_email:
        return {"conflict": False}

    requested_start, requested_end = build_event_datetimes(
        date, start_time, end_time
    )
    duration = requested_end - requested_start #how long requested meeting lasts
#bdecides how much of the calendar to search
    day_start = requested_start.replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    latest_start = day_start.replace(hour=LATEST_SUGGESTION_HOUR)

    query_end = max(
        day_start + timedelta(days=1), latest_start + duration, requested_end
    )

    ignored = str(ignore_event_id or "").strip()
    busy = []
#goes through all outlook
    for event in get_calendar_view(user_email, day_start, query_end):
        # The event we are moving must not block its own new time.
        if ignored and str(event.get("id", "") or "").strip() == ignored:
            continue

        if event.get("isCancelled"):
            continue

        show_as = str(event.get("showAs", "")).strip().lower()

        if show_as == "free":
            continue
#for each busy time
        event_start = parse_calendar_datetime(
            event.get("start", {}).get("dateTime", "")
        )
        event_end = parse_calendar_datetime(
            event.get("end", {}).get("dateTime", "")
        )

        if not event_start or not event_end:
            continue

        title = event.get("subject", "No title")
        busy.append((event_start, event_end, title))

        log.info(
            "Calendar busy event: %s | %s - %s | showAs=%s",
            title,
            event_start.strftime("%Y-%m-%d %I:%M %p"),
            event_end.strftime("%Y-%m-%d %I:%M %p"),
            show_as or "unknown",
        )
#requested against busy is compared
    conflicting_title = next(
        (
            title
            for start, end, title in busy
            if times_overlap(requested_start, requested_end, start, end)
        ),
        "",
    )

    if not conflicting_title:
        return {
            "conflict": False,
            "requested_start": start_time,
            "requested_end": end_time,
        }

    now = datetime.now(TIMEZONE)
    candidate_start = requested_start

    while True:
        candidate_start += timedelta(minutes=SLOT_STEP_MINUTES)

        if candidate_start > latest_start:
            break

        # Do not suggest a slot that has already gone by today.
        if candidate_start.date() == now.date() and candidate_start < now:
            continue

        candidate_end = candidate_start + duration
        clashes = [
            title
            for start, end, title in busy
            if times_overlap(candidate_start, candidate_end, start, end)
        ]

        log.info(
            "Checking suggested slot: %s - %s | conflicts=%s",
            candidate_start.strftime("%I:%M %p"),
            candidate_end.strftime("%I:%M %p"),
            clashes,
        )

        if not clashes:
            log.info(
                "Free calendar slot found: %s - %s",
                candidate_start.strftime("%I:%M %p"),
                candidate_end.strftime("%I:%M %p"),
            )
            return {
                "conflict": True,
                "title": conflicting_title,
                "suggested_start": candidate_start.strftime("%H:%M"),
                "suggested_end": candidate_end.strftime("%H:%M"),
            }

    log.info("No free slot found before %02d:00.", LATEST_SUGGESTION_HOUR)

    return {
        "conflict": True,
        "title": conflicting_title,
        "suggested_start": None,
        "suggested_end": None,
    }

# retrieving and displaying a group of events for a certain day or upcoming period.
def get_calendar_events(user_email, day="today"):
    if not user_email:
        return "Calendar email is not configured."

    now = datetime.now(TIMEZONE)
    show_upcoming = day == "upcoming"

    if show_upcoming:
        start_datetime, end_datetime = now, now + timedelta(days=30)
        label = "the next 30 days"
    else:
        if day == "today":
            target_date = now.date()
        elif day == "tomorrow":
            target_date = (now + timedelta(days=1)).date()
        else:
            target_date = datetime.strptime(day, "%Y-%m-%d").date()

        start_datetime = datetime(
            target_date.year,
            target_date.month,
            target_date.day,
            tzinfo=TIMEZONE,
        )
        end_datetime = start_datetime + timedelta(days=1)
        label = day

    lines = []

    for event in get_calendar_view(user_email, start_datetime, end_datetime):
        start_dt = parse_calendar_datetime(
            event.get("start", {}).get("dateTime", "")
        )

        if not start_dt:
            continue

        time_format = "%b %d, %I:%M %p" if show_upcoming else "%I:%M %p"
        lines.append(
            f"{start_dt.strftime(time_format)} - "
            f"{event.get('subject', 'No title')}"
        )

    if not lines:
        return f"You do not have any events for {label}."

    return "\n".join([f"Your schedule for {label}:"] + lines)

# back up
def find_event_id(user_email, title, date, created_at=""):
    if not user_email or not title or not date:
        return ""

    day_start = datetime.strptime(date, "%Y-%m-%d").replace(tzinfo=TIMEZONE)
    target = title.strip().casefold()

    matches = [
        event
        for event in get_calendar_view(
            user_email, day_start, day_start + timedelta(days=1)
        )
        if event.get("subject", "").strip().casefold() == target
    ]

    if not matches:
        return ""

    if len(matches) == 1:
        return matches[0].get("id", "")

    task_created = parse_graph_time(created_at)

    if not task_created:
        return ""

    best_event = None
    best_gap = None

    for event in matches:
        event_created = parse_graph_time(event.get("createdDateTime", ""))

        if not event_created:
            continue

        gap = abs((event_created - task_created).total_seconds())

        if best_gap is None or gap < best_gap:
            best_event, best_gap = event, gap

    if best_gap is None or best_gap > 300:
        return ""

    return best_event.get("id", "")
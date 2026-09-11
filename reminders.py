import logging
import os
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from sharepoint import normalize_number
from calendar_client import (
    TIMEZONE,
    CALENDAR_TIMEZONE,
    CALENDAR_USER_EMAIL,
    get_calendar_view,
    parse_calendar_datetime,
)
from whatsapp import send_button_message

log = logging.getLogger(__name__)

REMINDER_LEAD_MINUTES = 15
SNOOZE_MINUTES = 5
REMINDER_RECIPIENT = os.getenv("REMINDER_WHATSAPP_NUMBER", "").strip()

_reminder_state = {}


def _send_reminder(event_id, title, start_dt):
    send_button_message(
        REMINDER_RECIPIENT,
        f'Reminder: "{title}" starts at {start_dt.strftime("%I:%M %p")}.',
        [
            (f"snooze:{event_id}", f"Snooze {SNOOZE_MINUTES} min"),
            (f"dismiss:{event_id}", "Dismiss"),
        ],
    )

def check_reminders():
    if not REMINDER_RECIPIENT or not CALENDAR_USER_EMAIL:
        return

    now = datetime.now(TIMEZONE)

    try:
        events = get_calendar_view(
            CALENDAR_USER_EMAIL,
            now,
            now + timedelta(minutes=REMINDER_LEAD_MINUTES),
        )
    except Exception:
        log.exception("Could not check the calendar for reminders.")
        return

    seen_ids = set()

    for event in events:
        event_id = event.get("id")
        start_dt = parse_calendar_datetime(
            event.get("start", {}).get("dateTime", "")
        )

        # Already started, so there is nothing left to warn about.
        if not event_id or not start_dt or start_dt < now:
            continue

        seen_ids.add(event_id)

        state = _reminder_state.setdefault(
            event_id,
            {
                "trigger_at": start_dt
                - timedelta(minutes=REMINDER_LEAD_MINUTES),
                "dismissed": False,
            },
        )

        if state["dismissed"] or now < state["trigger_at"]:
            continue

        try:
            _send_reminder(event_id, event.get("subject", "No title"), start_dt)
        except Exception:
            log.exception("Could not send reminder for event %s", event_id)

        state["trigger_at"] = start_dt + timedelta(hours=1)

    for old_id in list(_reminder_state):
        if old_id not in seen_ids:
            _reminder_state.pop(old_id, None)

def handle_reminder_action(action, event_id, sender):
    if normalize_number(sender) != normalize_number(REMINDER_RECIPIENT):
        return None

    state = _reminder_state.get(event_id)

    if not state:
        return None

    if action == "dismiss":
        state["dismissed"] = True
        return "Reminder dismissed."

    if action == "snooze":
        state["trigger_at"] = datetime.now(TIMEZONE) + timedelta(
            minutes=SNOOZE_MINUTES
        )
        state["dismissed"] = False
        return f"Okay, I will remind you again in {SNOOZE_MINUTES} minutes."

    return None

def start_reminder_scheduler():
    scheduler = BackgroundScheduler(timezone=CALENDAR_TIMEZONE)
    scheduler.add_job(check_reminders, "interval", seconds=60)
    scheduler.start()
    log.info("Reminder scheduler started.")

    return scheduler
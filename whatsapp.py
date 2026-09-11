import logging
import os
import certifi
import requests
from dotenv import load_dotenv

load_dotenv()

VERIFY_TOKEN = os.getenv(
    "WHATSAPP_VERIFY_TOKEN",
    ""
).strip()

PHONE_NUMBER_ID = os.getenv(
    "WHATSAPP_PHONE_NUMBER_ID",
    ""
).strip()

ACCESS_TOKEN = os.getenv(
    "WHATSAPP_ACCESS_TOKEN",
    ""
).strip()

GRAPH_API_VERSION = "v23.0"

GRAPH_BASE_URL = (
    f"https://graph.facebook.com/"
    f"{GRAPH_API_VERSION}"
)

MESSAGE_URL = (
    f"{GRAPH_BASE_URL}/"
    f"{PHONE_NUMBER_ID}/messages"
)

MEDIA_URL = (
    f"{GRAPH_BASE_URL}/"
    f"{PHONE_NUMBER_ID}/media"
)

log = logging.getLogger(__name__)


# Reuse HTTP connection
session = requests.Session()


# Say plainly what Meta rejected, instead of leaving a bare
# status code in a traceback
def log_failed_response(
    response,
    *args,
    **kwargs
):
    if response.status_code == 401:
        log.error(
            "WhatsApp rejected the access token (401). "
            "WHATSAPP_ACCESS_TOKEN in .env has expired or "
            "been revoked - generate a new one."
        )
    elif not response.ok:
        log.error(
            "WhatsApp %s -> %s %s",
            response.url,
            response.status_code,
            response.text[:1000],
        )

    return response


session.hooks["response"].append(
    log_failed_response
)

AUTH_HEADERS = {
    "Authorization":
        f"Bearer {ACCESS_TOKEN}"
}

# Send text message
def send_message(
    recipient,
    body,
):
    response = session.post(
        MESSAGE_URL,
        headers={
            **AUTH_HEADERS,
            "Content-Type":
                "application/json",
        },
        json={
            "messaging_product":
                "whatsapp",
            "to": recipient,
            "type": "text",
            "text": {
                "body": body
            },
        },
        timeout=20,
        verify=certifi.where(),
    )
    response.raise_for_status()
    return True

# Send an interactive message with up to 3 reply buttons
def send_button_message(
    recipient,
    body,
    buttons,
):
    response = session.post(
        MESSAGE_URL,
        headers={
            **AUTH_HEADERS,
            "Content-Type":
                "application/json",
        },
        json={
            "messaging_product":
                "whatsapp",
            "to": recipient,
            "type": "interactive",
            "interactive": {
                "type": "button",
                "body": {
                    "text": body
                },
                "action": {
                    "buttons": [
                        {
                            "type": "reply",
                            "reply": {
                                "id": button_id,
                                "title": button_title,
                            },
                        }
                        for button_id, button_title in buttons
                    ]
                },
            },
        },
        timeout=20,
        verify=certifi.where(),
    )
    response.raise_for_status()
    return True

# Send Word document
def send_document(
    recipient,
    filename,
):
    with open(
        filename,
        "rb"
    ) as file:
        response = session.post(
            MEDIA_URL,
            headers=AUTH_HEADERS,
            files={
                "file": (
                    os.path.basename(
                        filename
                    ),
                    file,
                    "application/vnd."
                    "openxmlformats-officedocument."
                    "wordprocessingml.document",
                )
            },
            data={
                "messaging_product":
                    "whatsapp"
            },
            timeout=30,
            verify=certifi.where(),
        )
    response.raise_for_status()
    media_id = response.json()["id"]

    response = session.post(
        MESSAGE_URL,
        headers={
            **AUTH_HEADERS,
            "Content-Type":
                "application/json",
        },
        json={
            "messaging_product":
                "whatsapp",
            "to": recipient,
            "type": "document",
            "document": {
                "id": media_id,
                "filename":
                    os.path.basename(
                        filename
                    ),
            },
        },
        timeout=20,
        verify=certifi.where(),
    )
    response.raise_for_status()
    return True

# Common media download helper
def download_media(
    media_id
):
    if not media_id:
        raise ValueError(
            "WhatsApp media ID is missing."
        )
    if not ACCESS_TOKEN:
        raise ValueError(
            "WHATSAPP_ACCESS_TOKEN is missing."
        )

    # Get temporary media URL
    response = session.get(
        f"{GRAPH_BASE_URL}/{media_id}",
        headers=AUTH_HEADERS,
        timeout=20,
        verify=certifi.where(),
    )
    response.raise_for_status()
    media_info = response.json()
    media_url = media_info.get(
        "url"
    )

    if not media_url:
        raise RuntimeError(
            "Meta did not return a "
            "media download URL."
        )

    # Download actual media
    response = session.get(
        media_url,
        headers=AUTH_HEADERS,
        timeout=60,
        verify=certifi.where(),
    )
    response.raise_for_status()
    return (
        response.content,
        media_info.get(
            "mime_type",
            "application/octet-stream",
        ),
    )

# Download document
def download_document(
    media_id
):
    content, _ = download_media(
        media_id
    )
    return content

# Download voice/audio
def download_audio(
    media_id
):
    return download_media(
        media_id
    )
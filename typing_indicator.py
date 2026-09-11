import os
import certifi
import requests
from dotenv import load_dotenv

#loads Whatsapp information to communicate with whatsapp cloud API
load_dotenv()
PHONE_NUMBER_ID = os.getenv("WHATSAPP_PHONE_NUMBER_ID", "").strip()
ACCESS_TOKEN = os.getenv("WHATSAPP_ACCESS_TOKEN", "").strip()

#creating meta graph API 
GRAPH_API_VERSION = "v23.0"

GRAPH_URL = (
    f"https://graph.facebook.com/"
    f"{GRAPH_API_VERSION}/{PHONE_NUMBER_ID}/messages"
)


def send_typing_indicator(message_id):
 #making sure eveyrthing is set up correctly 
    if not ACCESS_TOKEN:
        return False

    if not PHONE_NUMBER_ID:
        return False

    if not message_id:
        return False

#authentication headers 
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json",
    }

#creating type indicator 
    payload = {
        "messaging_product": "whatsapp",
        "status": "read",
        "message_id": message_id,
        "typing_indicator": {
            "type": "text"
        },
    }

    try: #sends HTTP post reuqest 
        response = requests.post(
            GRAPH_URL,
            headers=headers,
            json=payload,
            timeout=10,
            verify=certifi.where(),
        )

        return response.ok

    except requests.RequestException:
        return False
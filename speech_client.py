import os
import certifi
import requests
from dotenv import load_dotenv

load_dotenv()

SPEECH_KEY = os.getenv("AZURE_SPEECH_KEY", "").strip()
SPEECH_REGION = os.getenv("AZURE_SPEECH_REGION", "").strip()

session = requests.Session()

# Send a WhatsApp voice note straight to Azure Speech and get the text back.
# WhatsApp gives us OGG/Opus, which the REST endpoint accepts as-is.
def transcribe_audio(audio_bytes, mime_type="audio/ogg", language="en-US"):
    if not SPEECH_KEY:
        raise ValueError("AZURE_SPEECH_KEY is missing.")
    if not SPEECH_REGION:
        raise ValueError("AZURE_SPEECH_REGION is missing.")

    response = session.post(
        f"https://{SPEECH_REGION}.stt.speech.microsoft.com"
        f"/speech/recognition/conversation/cognitiveservices/v1",
        params={"language": language, "format": "detailed"},
        headers={
            "Ocp-Apim-Subscription-Key": SPEECH_KEY,
            "Content-Type": mime_type,
            "Accept": "application/json",
        },
        data=audio_bytes,
        timeout=60,
        verify=certifi.where(),
    )
    response.raise_for_status()
    result = response.json()

    # DisplayText is the punctuated version and is the one we want. Some
    # responses leave it empty and only fill in NBest, so fall back to that.
    text = result.get("DisplayText", "").strip()

    nbest = result.get("NBest", [])
    if not text and nbest:
        text = nbest[0].get("Display", "").strip()

    if not text:
        # Nothing recognised usually means silence or a language mismatch.
        raise RuntimeError("Azure Speech could not recognize the audio.")

    return text

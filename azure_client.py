import os
import json
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from openai import AzureOpenAI, BadRequestError

load_dotenv()

AZURE_AI_ENDPOINT = os.getenv("AZURE_AI_ENDPOINT", "").strip()
AZURE_AI_API_KEY = os.getenv("AZURE_AI_API_KEY", "").strip()
AZURE_AI_DEPLOYMENT = os.getenv("AZURE_AI_DEPLOYMENT", "").strip()
AZURE_AI_API_VERSION = os.getenv("AZURE_AI_API_VERSION", "").strip()
AZURE_AI_EMBEDDING_DEPLOYMENT = os.getenv(
    "AZURE_AI_EMBEDDING_DEPLOYMENT", ""
).strip()

BEIRUT = ZoneInfo("Asia/Beirut")

client = AzureOpenAI(
    api_key=AZURE_AI_API_KEY,
    api_version=AZURE_AI_API_VERSION,
    azure_endpoint=AZURE_AI_ENDPOINT,
)

class ContentFilterError(Exception):
    pass

def is_content_filter(error):
    body = getattr(error, "body", None)

    if isinstance(body, dict):
        inner = body.get("error")
        if body.get("code") == "content_filter":
            return True
        if isinstance(inner, dict) and inner.get("code") == "content_filter":
            return True

    return "content_filter" in str(error)

def get_response(messages, **options):
    try:
        response = client.chat.completions.create(
            model=AZURE_AI_DEPLOYMENT,
            messages=messages,
            **options,
        )
    except BadRequestError as error:
        if is_content_filter(error):
            raise ContentFilterError(str(error)) from error
        raise

    return response.choices[0].message.content or ""

def get_json_response(system_message, text):
    response = get_response([
        {"role": "system", "content": system_message},
        {"role": "user", "content": text},
    ])

    return json.loads(
        response.replace("```json", "").replace("```", "").strip()
    )

# Turns a piece of text into a vector so it can be compared by meaning.
def embed(text):
    if not AZURE_AI_EMBEDDING_DEPLOYMENT:
        raise RuntimeError(
            "AZURE_AI_EMBEDDING_DEPLOYMENT is not set in the .env file."
        )

    response = client.embeddings.create(
        model=AZURE_AI_EMBEDDING_DEPLOYMENT,
        input=text,
    )

    return response.data[0].embedding

def ask_azure(
    message_text,
    history=None,
    sharepoint_context="",
    recalled_context="",
):
    now = datetime.now(BEIRUT)

    system_message = (
        "You are a helpful WhatsApp AI assistant. "
        "Keep responses concise, clear, and friendly. "
        f"Today's date is {now.strftime('%A, %B %d, %Y')}. "
        f"The current time in Beirut is {now.strftime('%I:%M %p')}."
    )

    if sharepoint_context:
        system_message += (
            "\n\nUse this SharePoint company information when relevant. "
            "Do not invent company information.\n\n"
            + sharepoint_context
        )

    if recalled_context:
        system_message += (
            "\n\nThese are things the user told you earlier in older "
            "conversations. Use them only when they are relevant to the "
            "current question, and never mention that you looked them "
            "up.\n\n"
            + recalled_context
        )

    return get_response([
        {"role": "system", "content": system_message},
        *(history or []),
        {"role": "user", "content": message_text},
    ])

def generate_document_content(request_text, sharepoint_context=""):
    system_message = (
        "Create a clear professional business document. "
        "Use SharePoint information when the request is company-related. "
        "Do not invent company information."
    )

    if sharepoint_context:
        system_message += "\n\nSharePoint Information:\n" + sharepoint_context

    return get_response([
        {"role": "system", "content": system_message},
        {"role": "user", "content": request_text},
    ])

def summarize_document(document_text):
    if not document_text or not document_text.strip():
        raise ValueError("The document contains no readable text.")

    return get_response([
        {
            "role": "system",
            "content": (
                "You are a document summarization assistant. "
                "Summarize ONLY the document text provided by the user. "
                "Include the main purpose, important points, "
                "and any actions or deadlines."
            ),
        },
        {
            "role": "user",
            "content": (
                "Here is the document text to summarize:\n\n" + document_text
            ),
        },
    ])

def extract_task_details(text):
    now = datetime.now(BEIRUT)

    return get_json_response(
        "Extract task information from the user's message. "
        f"Today's date is {now.strftime('%Y-%m-%d')} "
        f"({now.strftime('%A')}). Use this to resolve "
        "words like 'today' or 'tomorrow'. "
        "Return ONLY valid JSON with these keys: "
        "task, due, priority. "
        "The 'task' field must be a SHORT descriptive "
        "name for the task, not the entire message and "
        "not the date/time/priority wording. For example, "
        "if the message is 'Book a meeting called Test "
        "Conflict at 12pm to 1:30pm today', the task "
        "should be 'Test Conflict', not the full sentence. "
        "Use YYYY-MM-DD for due. If no due date is "
        "mentioned, return an empty string for due. "
        "Priority must be High, Medium, or Low.",
        text,
    )

def extract_expense_details(text):
    return get_json_response(
        "Extract expense information from the "
        "user's message. Return ONLY valid JSON "
        "with these keys: description, amount, "
        "category. "
        "The 'description' field must be a SHORT "
        "name for what was purchased (e.g. 'lunch'), "
        "not the entire message. "
        "The 'amount' field must be a plain number "
        "with no currency symbol, e.g. 20 or 20.5. "
        "Always assume the amount is in US dollars, "
        "even if another currency is mentioned in "
        "the message. "
        "The 'category' field must be one of: Food, "
        "Transport, Shopping, Entertainment, Bills, "
        "Other.",
        text,
    )

def extract_budget_details(text):
    return get_json_response(
        "Extract a monthly budget limit from the "
        "user's message. Return ONLY valid JSON "
        "with these keys: category, amount. "
        "The 'category' field must be one of: Food, "
        "Transport, Shopping, Entertainment, Bills, "
        "Other - map everyday words to the closest "
        "one, e.g. 'taxi' or 'uber' maps to Transport, "
        "'groceries' maps to Food, 'movies' maps to "
        "Entertainment. "
        "The 'amount' field must be a plain number "
        "with no currency symbol, e.g. 500 or 150.5.",
        text,
    )

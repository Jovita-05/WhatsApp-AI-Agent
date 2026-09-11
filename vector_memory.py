import json
import logging
import math
import os
from azure_client import embed

log = logging.getLogger(__name__)

MAX_HISTORY_MESSAGES = 10
STORE_PATH = "vector_memory.json"
TOP_K = 3
SIMILARITY_THRESHOLD = 0.30
MIN_TEXT_LENGTH = 15
RECENT_WINDOW = 6
conversation_memory = {}
_memories = {}

#old memory.py (recent memory)
def get_history(user_id):
    return conversation_memory.get(user_id, [])

def _append(user_id, role, message):
    history = conversation_memory.setdefault(user_id, [])
    history.append(
        {
            "role": role,
            "content": message,
        }
    )
    if len(history) > MAX_HISTORY_MESSAGES:
        conversation_memory[user_id] = history[-MAX_HISTORY_MESSAGES:]

#new memory 
#bridge between both memories
def add_user_message(user_id, message):
    _append(user_id, "user", message)
    remember(user_id, message)

def add_assistant_message(user_id, message):
    _append(user_id, "assistant", message)

def clear_history(user_id):
    conversation_memory.pop(user_id, None)
    if _memories.pop(user_id, None) is not None:
        _save()

#loads the saved long term vector memories
def _load():
    if not os.path.exists(STORE_PATH):
        return
    try:
        with open(STORE_PATH, encoding="utf-8") as store_file:
            _memories.update(json.load(store_file))
    except Exception:
        log.exception("Could not read the memory store, starting empty.")

def _save():
    try:
        with open(STORE_PATH, "w", encoding="utf-8") as store_file:
            json.dump(_memories, store_file)
    except Exception:
        log.exception("Could not save the memory store.")

#comapres two embedding vectors
def cosine_similarity(a, b):
    dot = sum(x * y for x, y in zip(a, b))
    norm_a = math.sqrt(sum(x * x for x in a))
    norm_b = math.sqrt(sum(y * y for y in b))

    return dot / (norm_a * norm_b) if norm_a and norm_b else 0.0

#prepares user message for long-term
def remember(user_id, text):
    if len(text.strip()) < MIN_TEXT_LENGTH:
        return
    try:
        vector = embed(text)
    except Exception:
        log.exception("Could not embed a message for long term memory.")
        return
    _memories.setdefault(user_id, []).append([text, vector])
    _save()

#find older memory
def recall(user_id, query):
    stored = _memories.get(user_id, [])

    candidates = stored[:-RECENT_WINDOW] if len(stored) > RECENT_WINDOW else []

    if not candidates:
        return []

    query_vector = embed(query)

    scored = sorted(
        (
            (cosine_similarity(query_vector, vector), text)
            for text, vector in candidates
        ),
        reverse=True,
    )

    return [
        text for score, text in scored[:TOP_K] if score >= SIMILARITY_THRESHOLD
    ]

def recall_context(user_id, query):
    try:
        return "\n".join("- " + text for text in recall(user_id, query))
    except Exception:
        log.exception("Long term memory recall failed.")
        return ""

_load()

from datetime import datetime
from zoneinfo import ZoneInfo

from vector_memory import clear_history


def handle_command(sender, text_body):
    command = text_body.lower().strip()

    if command == "/help":
        return (
            "🤖 WhatsApp AI Assistant\n\n"
            "Available commands:\n"
            "/help - Show available commands\n"
            "/about - About this assistant\n"
            "/clear - Clear conversation memory\n"
            "/date - Show today's date\n"
            "/time - Show the current time\n\n"
            "Or simply send me any question."
        )

    if command == "/about":
        return (
            "🤖 Azure AI WhatsApp Assistant\n\n"
            "Built with:\n"
            "• Python\n"
            "• Flask\n"
            "• Azure AI Foundry\n"
            "• WhatsApp Cloud API\n\n"
            "Features:\n"
            "• Conversation memory\n"
            "• Date and time awareness\n"
            "• AI-powered responses"
        )

    if command == "/clear":
        clear_history(sender)
        return "✅ Your conversation history has been cleared."

    if command == "/date":
        now = datetime.now(ZoneInfo("Asia/Beirut"))
        return now.strftime("Today is %A, %B %d, %Y.")

    if command == "/time":
        now = datetime.now(ZoneInfo("Asia/Beirut"))
        return now.strftime(
            "The current time in Beirut is %I:%M %p."
        )

    return None
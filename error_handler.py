import logging
from azure_client import ContentFilterError

# create logger to record errors 
log = logging.getLogger(__name__)

#unexpected error 
def handle_error(error):

    # A blocked wording is not a bug, so log one line about it
    # instead of a traceback and tell the user what to do
    if isinstance(error, ContentFilterError):
        log.warning(
            "Azure content filter rejected the message: %s",
            error,
        )

        return (
            "⚠️ I couldn't process the wording of that message.\n"
            "Please try rephrasing it."
        )

    # The caller already logged the traceback, so keep this to
    # one line instead of repeating the whole stack
    log.error("Application error: %s", error)

    return (
        "⚠️ Sorry, I couldn't process your request right now.\n"
        "Please try again in a few moments."
    )

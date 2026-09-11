import sharepoint

REQUEST_LIST = "WhatsApp Requests"

# Add request to SharePoint
def add_request(
    name,
    department,
    request_text,
    whatsapp_number,
):
    headers, site_id, list_id = (
        sharepoint.get_list_context(
            REQUEST_LIST
        )
    )

    response = sharepoint.session.post(
        f"{sharepoint.GRAPH_URL}/sites/"
        f"{site_id}/lists/"
        f"{list_id}/items",
        headers={
            **headers,
            "Content-Type":
                "application/json",
        },
        json={
            "fields": {
                "Title": name,
                "Department": department,
                "Request": request_text,
                "WhatsAppNumber":
                    str(whatsapp_number),
                "Status": "Pending",
            }
        },
        timeout=30,
    )

    response.raise_for_status()

    return response.json()

# Get request status
def get_request_status(
    whatsapp_number,
):
    try:
        headers, site_id, list_id = (
            sharepoint.get_list_context(
                REQUEST_LIST
            )
        )

    except RuntimeError:
        return (
            "WhatsApp Requests list "
            "was not found."
        )

    response = sharepoint.session.get(
        f"{sharepoint.GRAPH_URL}/sites/"
        f"{site_id}/lists/"
        f"{list_id}/items"
        f"?expand=fields",
        headers=headers,
        timeout=30,
    )

    response.raise_for_status()

    number = str(
        whatsapp_number
    )

    for item in (
        response.json().get(
            "value",
            [],
        )
    ):
        fields = item.get(
            "fields",
            {},
        )

        if (
            str(
                fields.get(
                    "WhatsAppNumber",
                    "",
                )
            )
            == number
        ):
            return (
                f"Your request: "
                f"{fields.get('Request', '')}\n"
                f"Status: "
                f"{fields.get('Status', 'Pending')}"
            )

    return (
        "I could not find a request "
        "for your number."
    )
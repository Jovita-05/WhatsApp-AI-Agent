from datetime import datetime
import sharepoint
from sharepoint import json_headers, list_items_url
from calendar_client import TIMEZONE, to_local_date

EXPENSE_LIST = "WhatsApp Expenses"
BUDGET_LIST = "WhatsApp Budgets"

def _all_items(list_name):
    headers, site_id, list_id = sharepoint.get_list_context(list_name)

    response = sharepoint.session.get(
        list_items_url(site_id, list_id) + "?expand=fields",
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()

    return response.json().get("value", [])


def _owned_by(items, whatsapp_number):
    wanted = str(whatsapp_number)

    return [
        item
        for item in items
        if str(item.get("fields", {}).get("WhatsAppNumber", "")) == wanted
    ]


def add_expense(description, amount, category, whatsapp_number):
    headers, site_id, list_id = sharepoint.get_list_context(EXPENSE_LIST)

    response = sharepoint.session.post(
        list_items_url(site_id, list_id),
        headers=json_headers(headers),
        json={
            "fields": {
                "Title": description,
                "Amount": amount,
                "Category": category,
                "Currency": "$",
                "ExpenseDate": datetime.now(TIMEZONE).strftime("%Y-%m-%d"),
                "WhatsAppNumber": str(whatsapp_number),
            }
        },
        timeout=30,
    )
    response.raise_for_status()

    return response.json()


def get_month_totals(whatsapp_number):
    month = datetime.now(TIMEZONE).strftime("%Y-%m")

    total = 0.0
    by_category = {}

    for item in _owned_by(_all_items(EXPENSE_LIST), whatsapp_number):
        fields = item.get("fields", {})

        if not to_local_date(fields.get("ExpenseDate", "")).startswith(month):
            continue

        amount = float(fields.get("Amount", 0) or 0)
        category = fields.get("Category", "Other") or "Other"

        total += amount
        by_category[category] = by_category.get(category, 0) + amount

    return total, by_category


def get_expenses_this_month(whatsapp_number):
    total, by_category = get_month_totals(whatsapp_number)

    if not by_category:
        return "You have no expenses recorded this month."

    return "\n".join(
        [f"Your expenses this month: ${total:.2f}"]
        + [
            f"- {category}: ${amount:.2f}"
            for category, amount in by_category.items()
        ]
    )

def get_budgets(whatsapp_number):
    return {
        item.get("fields", {}).get("Title", ""): {
            "id": item.get("id", ""),
            "amount": float(item.get("fields", {}).get("Amount", 0) or 0),
        }
        for item in _owned_by(_all_items(BUDGET_LIST), whatsapp_number)
    }


def set_budget(category, amount, whatsapp_number):
    headers, site_id, list_id = sharepoint.get_list_context(BUDGET_LIST)
    existing = get_budgets(whatsapp_number)

    # Patch when this category already has a budget, otherwise create one.
    if category in existing:
        response = sharepoint.session.patch(
            list_items_url(site_id, list_id, existing[category]["id"])
            + "/fields",
            headers=json_headers(headers),
            json={"Amount": amount},
            timeout=30,
        )
    else:
        response = sharepoint.session.post(
            list_items_url(site_id, list_id),
            headers=json_headers(headers),
            json={
                "fields": {
                    "Title": category,
                    "Amount": amount,
                    "WhatsAppNumber": str(whatsapp_number),
                }
            },
            timeout=30,
        )

    response.raise_for_status()

    return response.json()
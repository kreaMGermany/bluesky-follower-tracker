import os
import json
import requests
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"
BERLIN = ZoneInfo("Europe/Berlin")


def getenv(name):
    v = os.getenv(name)
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def get_token():
    app = ConfidentialClientApplication(
        getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
        client_credential=getenv("CLIENT_SECRET"),
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token["access_token"]


def get_followers(handle):
    url = f"https://public.api.bsky.app/xrpc/app.bsky.actor.getProfile?actor={handle}"
    r = requests.get(url)
    r.raise_for_status()
    return r.json()["followersCount"]


def graph_get(token, url):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    return r.json()


def graph_post(token, url, payload):
    r = requests.post(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=payload)
    r.raise_for_status()
    return r.json() if r.text else {}


def send_mail(token, sender, recipients, subject, html):
    to_recipients = [{"emailAddress": {"address": r.strip()}} for r in recipients.split(",")]

    graph_post(
        token,
        f"{GRAPH}/users/{sender}/sendMail",
        {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": html},
                "toRecipients": to_recipients
            }
        }
    )


def main():
    token = get_token()
    sender = getenv("SENDER_UPN")
    recipients = getenv("RECIPIENTS")
    path = getenv("ONEDRIVE_FILE_PATH")
    accounts = json.loads(getenv("ACCOUNTS_JSON"))

    today = datetime.now(BERLIN).date()

    # Excel laden
    file = graph_get(token, f"{GRAPH}/users/{sender}/drive/root:/{path}")
    file_id = file["id"]

    table = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables")["value"][0]["id"]

    rows_data = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables/{table}/rows?$top=5000")

    history = {}

    for r in rows_data["value"]:
        vals = r["values"][0]
        d = vals[0]
        acc = vals[1].replace("'", "")
        foll = int(vals[2])

        history.setdefault(acc, []).append((d, foll))

    results = []

    for acc in accounts:
        handle = acc["handle"]
        followers = get_followers(handle)

        prev = [x for x in history.get(handle, []) if x[0] < str(today)]
        base = prev[-1] if prev else (None, None)

        base_date, base_followers = base

        delta = followers - base_followers if base_followers else None
        delta_pct = (delta / base_followers * 100) if base_followers else None

        results.append({
            "account": handle,
            "followers": followers,
            "delta": delta,
            "delta_pct": delta_pct,
            "base_date": base_date
        })

    # HTML wie IG
    html = f"<h2>Bluesky Report {today}</h2>"

    html += "<h3>Alle Accounts</h3>"
    html += "<table border='1'><tr><th>Account</th><th>Followers</th><th>Δ</th><th>Δ %</th></tr>"

    for r in results:
        html += f"<tr><td>{r['account']}</td><td>{r['followers']}</td><td>{r['delta']}</td><td>{r['delta_pct']}</td></tr>"

    html += "</table>"

    send_mail(token, sender, recipients, f"Bluesky Report {today}", html)

    # speichern
    values = [[today.isoformat(), "'" + r["account"], r["followers"]] for r in results]

    graph_post(
        token,
        f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables/{table}/rows/add",
        {"values": values}
    )


if __name__ == "__main__":
    main()

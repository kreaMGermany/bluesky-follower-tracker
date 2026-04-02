import os
import json
import requests
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from typing import Any, Dict, List, Optional, Tuple
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"
BERLIN = ZoneInfo("Europe/Berlin")


def getenv(name: str) -> str:
    v = os.getenv(name)
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    app = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"Failed to get Graph token: {res}")
    return res["access_token"]


def graph_get(token: str, url: str) -> Dict[str, Any]:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()


def graph_post(token: str, url: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=json.dumps(payload),
        timeout=60,
    )
    r.raise_for_status()
    return r.json() if r.text else {}


def parse_excel_date(x: Any) -> Optional[date]:
    if x is None:
        return None
    if isinstance(x, (int, float)):
        base = date(1899, 12, 30)
        return base + timedelta(days=int(float(x)))

    s = str(x).strip()
    if not s:
        return None
    if "T" in s:
        s = s[:10]

    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except:
        try:
            return datetime.strptime(s[:10], "%d.%m.%Y").date()
        except:
            return None


def normalize_account(x: Any) -> str:
    return str(x).lstrip("'").strip()


def fetch_bluesky_followers(handle: str):
    url = f"https://public.api.bsky.app/xrpc/app.bsky.actor.getProfile?actor={handle}"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    data = r.json()
    return data.get("handle"), int(data.get("followersCount", 0))


def send_mail(token: str, sender: str, recipients: str, subject: str, html: str):
    to_recipients = [{"emailAddress": {"address": r.strip()}} for r in recipients.split(",")]
    url = f"{GRAPH}/users/{sender}/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": to_recipients,
        }
    }
    graph_post(token, url, payload)


def main():
    tenant_id = getenv("TENANT_ID")
    client_id = getenv("CLIENT_ID")
    client_secret = getenv("CLIENT_SECRET")
    sender_upn = getenv("SENDER_UPN")
    recipients = getenv("RECIPIENTS")
    onedrive_path = getenv("ONEDRIVE_FILE_PATH")
    accounts = json.loads(getenv("ACCOUNTS_JSON"))

    today = datetime.now(BERLIN).date()
    token = graph_token(tenant_id, client_id, client_secret)

    file_meta = graph_get(token, f"{GRAPH}/users/{sender_upn}/drive/root:/{onedrive_path}")
    file_id = file_meta["id"]

    table = graph_get(token, f"{GRAPH}/users/{sender_upn}/drive/items/{file_id}/workbook/tables")["value"][0]["id"]

    rows_data = graph_get(token, f"{GRAPH}/users/{sender_upn}/drive/items/{file_id}/workbook/tables/{table}/rows?$top=5000")

    history = {}
    for r in rows_data["value"]:
        vals = r["values"][0]
        d = parse_excel_date(vals[0])
        acc = normalize_account(vals[1])
        foll = int(vals[2])
        history.setdefault(acc, []).append((d, foll))

    results = []

    for acc in accounts:
        handle = acc["handle"]
        _, followers = fetch_bluesky_followers(handle)

        prev = [(d, f) for (d, f) in history.get(handle, []) if d < today]
        base_date, base_followers = (prev[-1] if prev else (None, None))

        delta = followers - base_followers if base_followers else None
        delta_pct = (delta / base_followers * 100) if base_followers else None

        results.append((handle, followers, delta, delta_pct))

    html = "<h2>Bluesky Report</h2><table border='1'><tr><th>Account</th><th>Followers</th><th>Δ</th><th>Δ %</th></tr>"

    for r in results:
        html += f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td></tr>"

    html += "</table>"

    send_mail(token, sender_upn, recipients, f"Bluesky Report {today}", html)

    values = [[today.isoformat(), "'" + r[0], r[1]] for r in results]

    graph_post(token, f"{GRAPH}/users/{sender_upn}/drive/items/{file_id}/workbook/tables/{table}/rows/add", {"values": values})


if __name__ == "__main__":
    main()

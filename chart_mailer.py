import os
import json
import base64
import requests
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from datetime import datetime, date, timedelta
from collections import defaultdict
from PIL import Image
from msal import ConfidentialClientApplication
import tempfile

GRAPH = "https://graph.microsoft.com/v1.0"
BG = '#E8E8E8'
ACCENT = '#1a1a1a'
TEXT = '#111111'
TEXT2 = '#555555'
GRID = '#cccccc'
GREEN = '#1e7d1e'

COLORS = [
    '#1a1a1a', '#0077C8', '#CC3399', '#3DAA3D',
    '#7B0DAA', '#C8870A', '#AA1A1A', '#1A7BAA', '#5A5A5A'
]


def getenv(name):
    v = os.getenv(name)
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def graph_get(token, url):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    r.raise_for_status()
    return r.json()


def graph_post(token, url, payload):
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    r.raise_for_status()
    return r.json() if r.text else {}


def get_token():
    app = ConfidentialClientApplication(
        getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
        client_credential=getenv("CLIENT_SECRET"),
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        raise RuntimeError(f"Failed to get Graph token: {token}")
    return token["access_token"]


def parse_excel_date(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(float(value)))
    s = str(value).strip()
    if not s:
        return None
    if "T" in s:
        s = s[:10]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def load_logo():
    """Load kreaM M-Icon logo."""
    logo_path = "kream_logo.jpg"
    if os.path.exists(logo_path):
        img = Image.open(logo_path).convert('RGB')
        return np.array(img.resize((90, 90), Image.LANCZOS))
    return None


def make_individual_chart(handle, name, day_data, logo_arr, tmpdir):
    """Generate individual model chart. Returns file path."""
    sorted_days = sorted(day_data.items())
    dates = [d for d, _ in sorted_days]
    followers = [f for _, f in sorted_days]
    start, end = followers[0], followers[-1]
    growth = end - start
    pct = growth / start * 100 if start > 0 else 0

    fig, ax = plt.subplots(figsize=(11, 5.5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)

    ax.fill_between(dates, followers, alpha=0.07, color=ACCENT)
    ax.plot(dates, followers, color=ACCENT, linewidth=2.5, zorder=3)
    ax.scatter([dates[-1]], [followers[-1]], color=ACCENT, s=70, zorder=4)
    ax.annotate(f'{end:,}'.replace(',', '.'),
        xy=(dates[-1], followers[-1]), xytext=(10, 0),
        textcoords='offset points', fontsize=11,
        fontweight='bold', color=ACCENT, va='center')

    fig.text(0.07, 0.90, name, color=TEXT, fontsize=21, fontweight='bold', transform=fig.transFigure)
    fig.text(0.07, 0.83, f'@{handle}', color=TEXT2, fontsize=10, transform=fig.transFigure)
    fig.text(0.62, 0.90, f'{end:,} Follower'.replace(',', '.'), color=TEXT, fontsize=14, fontweight='bold', transform=fig.transFigure)
    fig.text(0.62, 0.83, f'+{growth:,} Follower   +{pct:.0f}% seit Start'.replace(',', '.'), color=GREEN, fontsize=10, transform=fig.transFigure)
    period = f'{dates[0].strftime("%d.%m.%Y")} – {dates[-1].strftime("%d.%m.%Y")}'
    fig.text(0.07, 0.03, period, color=TEXT2, fontsize=8, transform=fig.transFigure)

    ax.tick_params(colors=TEXT, labelsize=10, length=3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d.%m.'))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
    plt.setp(ax.xaxis.get_majorticklabels(), color=TEXT)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', '.')))
    plt.setp(ax.yaxis.get_majorticklabels(), color=TEXT, fontsize=10)
    for spine in ax.spines.values():
        spine.set_edgecolor(GRID)
    ax.grid(axis='y', color=GRID, linewidth=0.8)
    ax.grid(axis='x', color=GRID, linewidth=0.5, linestyle='--')
    ax.set_ylim(bottom=0)

    if logo_arr is not None:
        newax = fig.add_axes([0.875, 0.80, 0.09, 0.16])
        newax.imshow(logo_arr)
        newax.axis('off')

    plt.tight_layout(rect=[0, 0.08, 1, 0.78])
    safe = handle.replace('.', '_').replace('-', '_')
    path = os.path.join(tmpdir, f'{safe}.png')
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor=BG)
    plt.close()
    return path


def make_overview_chart(all_data, display_names, logo_arr, tmpdir):
    """Generate overview chart with all models. Returns file path."""
    sorted_accounts = sorted(all_data.items(), key=lambda x: max(x[1].values()), reverse=True)

    all_dates_set = set()
    for _, dd in sorted_accounts:
        all_dates_set.update(dd.keys())
    all_dates = sorted(all_dates_set)

    fig, ax = plt.subplots(figsize=(14, 7))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)

    for i, (handle, day_data) in enumerate(sorted_accounts):
        sorted_days = sorted(day_data.items())
        dates = [d for d, _ in sorted_days]
        followers = [f for _, f in sorted_days]
        name = display_names.get(handle, handle)
        color = COLORS[i % len(COLORS)]
        end = followers[-1]

        ax.plot(dates, followers, color=color, linewidth=2.2, zorder=3,
                label=f'{name}  ({end:,})'.replace(',', '.'))
        ax.scatter([dates[-1]], [followers[-1]], color=color, s=55, zorder=4)
        ax.annotate(f'{end:,}'.replace(',', '.'),
            xy=(dates[-1], followers[-1]), xytext=(6, 0),
            textcoords='offset points', fontsize=8.5,
            fontweight='bold', color=color, va='center')

    today = datetime.utcnow().date()
    fig.text(0.05, 0.93, 'Bluesky Follower Übersicht', color=TEXT, fontsize=18, fontweight='bold', transform=fig.transFigure)
    period = f'{min(all_dates).strftime("%d.%m.%Y")} – {max(all_dates).strftime("%d.%m.%Y")}'
    fig.text(0.05, 0.87, period, color=TEXT2, fontsize=10, transform=fig.transFigure)

    ax.tick_params(colors=TEXT, labelsize=10)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d.%m.'))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
    plt.setp(ax.xaxis.get_majorticklabels(), color=TEXT)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', '.')))
    plt.setp(ax.yaxis.get_majorticklabels(), color=TEXT, fontsize=10)
    for spine in ax.spines.values():
        spine.set_edgecolor(GRID)
    ax.grid(axis='y', color=GRID, linewidth=0.8)
    ax.grid(axis='x', color=GRID, linewidth=0.5, linestyle='--')
    ax.set_ylim(bottom=0)
    ax.legend(loc='upper left', framealpha=0.85, facecolor=BG,
        edgecolor=GRID, fontsize=9.5, labelcolor=TEXT)

    if logo_arr is not None:
        newax = fig.add_axes([0.885, 0.87, 0.07, 0.12])
        newax.imshow(logo_arr)
        newax.axis('off')

    plt.tight_layout(rect=[0, 0.05, 1, 0.84])
    path = os.path.join(tmpdir, 'overview_alle_models.png')
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor=BG)
    plt.close()
    return path


def send_mail_with_attachment(token, sender, to_email, cc_emails, subject, html, attachment_path):
    """Send mail with PNG attachment via Microsoft Graph."""
    with open(attachment_path, 'rb') as f:
        content_bytes = base64.b64encode(f.read()).decode('utf-8')

    filename = os.path.basename(attachment_path)

    to_recipients = [{"emailAddress": {"address": to_email.strip()}}]
    cc_recipients = [
        {"emailAddress": {"address": r.strip()}}
        for r in cc_emails.split(",") if r.strip()
    ]

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentType": "image/png",
                    "contentBytes": content_bytes
                }
            ]
        },
        "saveToSentItems": True
    }
    graph_post(token, f"{GRAPH}/users/{sender}/sendMail", payload)


def build_model_html(name, handle, end_followers, growth, pct, period):
    return f"""
    <html>
      <body style="font-family:Segoe UI, Arial, sans-serif; color:#111; background:#f9f9f9; padding:24px;">
        <h2 style="margin:0 0 4px 0;">Dein wöchentlicher Bluesky Report</h2>
        <p style="color:#555; margin:0 0 20px 0;">{period}</p>
        <table style="border-collapse:collapse; width:400px;">
          <tr>
            <td style="padding:10px 16px; background:#E8E8E8; font-weight:600;">Account</td>
            <td style="padding:10px 16px; background:#f0f0f0;">@{handle}</td>
          </tr>
          <tr>
            <td style="padding:10px 16px; background:#E8E8E8; font-weight:600;">Follower</td>
            <td style="padding:10px 16px; background:#f0f0f0;">{end_followers:,}".replace(",",".")</td>
          </tr>
          <tr>
            <td style="padding:10px 16px; background:#E8E8E8; font-weight:600;">Wachstum</td>
            <td style="padding:10px 16px; background:#f0f0f0; color:#1e7d1e; font-weight:700;">+{growth:,} ({pct:.0f}%)".replace(",",".")</td>
          </tr>
        </table>
        <p style="margin:20px 0 4px 0; color:#555; font-size:13px;">Deine Wachstumskurve findest du im Anhang.</p>
        <p style="color:#555; font-size:12px; margin-top:24px;">kreaM Management</p>
      </body>
    </html>
    """.replace('".replace(",",".")', '')


def build_manager_html(today, all_data, display_names):
    rows = ""
    sorted_accounts = sorted(all_data.items(), key=lambda x: max(x[1].values()), reverse=True)
    for handle, day_data in sorted_accounts:
        sorted_days = sorted(day_data.items())
        followers = [f for _, f in sorted_days]
        start, end = followers[0], followers[-1]
        growth = end - start
        pct = growth / start * 100 if start > 0 else 0
        name = display_names.get(handle, handle)
        color = '#1e7d1e' if growth >= 0 else '#cc0000'
        rows += f"""
        <tr>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;font-weight:600;">{name}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;color:#555;font-size:12px;">@{handle}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;">{end:,}".replace(",",".")</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;color:{color};font-weight:700;">+{growth:,}".replace(",",".")</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;color:{color};font-weight:700;">+{pct:.0f}%</td>
        </tr>
        """
    return f"""
    <html>
      <body style="font-family:Segoe UI, Arial, sans-serif; color:#111; background:#f9f9f9; padding:24px;">
        <h2 style="margin:0 0 4px 0;">Bluesky Weekly Overview – {today.strftime("%d.%m.%Y")}</h2>
        <p style="color:#555; margin:0 0 20px 0;">Alle Models im Überblick. Wachstum seit Tracking-Start.</p>
        <table style="border-collapse:collapse; width:700px; max-width:100%;">
          <thead>
            <tr style="background:#E8E8E8;">
              <th style="text-align:left;padding:10px 8px;border-bottom:2px solid #333;">Name</th>
              <th style="text-align:left;padding:10px 8px;border-bottom:2px solid #333;">Handle</th>
              <th style="text-align:right;padding:10px 8px;border-bottom:2px solid #333;">Follower</th>
              <th style="text-align:right;padding:10px 8px;border-bottom:2px solid #333;">Δ</th>
              <th style="text-align:right;padding:10px 8px;border-bottom:2px solid #333;">Δ %</th>
            </tr>
          </thead>
          <tbody>{rows}</tbody>
        </table>
        <p style="color:#555; font-size:12px; margin-top:24px;">Übersichtschart im Anhang.</p>
      </body>
    </html>
    """.replace('".replace(",",".")', '')


def main():
    token = get_token()
    sender = getenv("SENDER_UPN")
    manager_emails = getenv("RECIPIENTS")
    path = getenv("ONEDRIVE_FILE_PATH")
    today = datetime.utcnow().date()

    # Load Excel
    file_meta = graph_get(token, f"{GRAPH}/users/{sender}/drive/root:/{path}")
    file_id = file_meta["id"]

    # Load Models tab (sheet name: "Models", columns: handle | name | email)
    sheets = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/worksheets")
    models_sheet_id = None
    log_sheet_id = None
    for s in sheets.get("value", []):
        if s["name"].lower() == "models":
            models_sheet_id = s["id"]
        if s["name"].lower() in ("log", "sheet1", "tabelle1"):
            log_sheet_id = s["id"]

    if not models_sheet_id:
        raise RuntimeError("Worksheet 'Models' not found in Excel")

    # Read models
    models_range = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/worksheets/{models_sheet_id}/usedRange")
    models = []
    display_names = {}
    for row in models_range.get("values", [])[1:]:  # skip header
        if len(row) < 3:
            continue
        handle = str(row[0]).strip()
        name = str(row[1]).strip()
        email = str(row[2]).strip()
        if handle and name and email:
            models.append({"handle": handle, "name": name, "email": email})
            display_names[handle] = name

    if not models:
        raise RuntimeError("No models found in Models sheet")

    print(f"Found {len(models)} models")

    # Read follower log
    tables = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables")
    if not tables.get("value"):
        raise RuntimeError("No Excel table found")
    table_id = tables["value"][0]["id"]
    rows_data = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/tables/{table_id}/rows?$top=5000")

    history = defaultdict(dict)
    for r in rows_data.get("value", []):
        vals = r.get("values", [[]])[0]
        if len(vals) < 3:
            continue
        d = parse_excel_date(vals[0])
        acc = str(vals[1]).replace("'", "").strip()
        try:
            foll = int(float(str(vals[2]).strip()))
        except Exception:
            continue
        if d and acc:
            history[acc][d] = foll

    # Load logo
    logo_arr = load_logo()

    with tempfile.TemporaryDirectory() as tmpdir:
        # Generate individual charts + send to each model
        for model in models:
            handle = model["handle"]
            name = model["name"]
            email = model["email"]

            if handle not in history or len(history[handle]) < 2:
                print(f"Skipping {handle} — not enough data")
                continue

            chart_path = make_individual_chart(handle, name, history[handle], logo_arr, tmpdir)

            day_data = history[handle]
            sorted_days = sorted(day_data.items())
            followers = [f for _, f in sorted_days]
            dates = [d for d, _ in sorted_days]
            start, end = followers[0], followers[-1]
            growth = end - start
            pct = growth / start * 100 if start > 0 else 0
            period = f'{dates[0].strftime("%d.%m.%Y")} – {dates[-1].strftime("%d.%m.%Y")}'

            html = build_model_html(name, handle, end, growth, pct, period)
            subject = f"Dein Bluesky Wachstum – {today.strftime('%d.%m.%Y')}"

            send_mail_with_attachment(token, sender, email, manager_emails, subject, html, chart_path)
            print(f"✓ Mail sent to {name} ({email})")

        # Generate overview + send to managers
        all_data = {m["handle"]: history[m["handle"]] for m in models if m["handle"] in history and len(history[m["handle"]]) >= 2}
        overview_path = make_overview_chart(all_data, display_names, logo_arr, tmpdir)

        manager_html = build_manager_html(today, all_data, display_names)
        # Send overview to first manager (others on CC) - simplest: TO = first, CC = rest
        manager_list = [r.strip() for r in manager_emails.split(",") if r.strip()]
        to_manager = manager_list[0]
        cc_managers = ",".join(manager_list[1:]) if len(manager_list) > 1 else ""

        send_mail_with_attachment(
            token, sender, to_manager, cc_managers,
            f"Bluesky Weekly Overview – {today.strftime('%d.%m.%Y')}",
            manager_html, overview_path
        )
        print(f"✓ Overview sent to managers")

    print("DONE")


if __name__ == "__main__":
    main()

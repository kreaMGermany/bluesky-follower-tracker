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
BG = '#F2F2F2'
# krea-M CI
CI_PRIMARY = '#005792'
CI_NAVY = '#00395F'
CI_AZUR = '#0094F8'
ACCENT = CI_NAVY
TEXT = '#111111'
TEXT2 = '#555555'
GRID = '#cfd6dc'
GREEN = '#1e7d1e'
RED = '#cc0000'
X_COLOR = CI_AZUR

COLORS = [
    '#00395F', '#005792', '#0075C5', '#0094F8',
    '#5A5A5A', '#7B0DAA', '#C8870A', '#AA1A1A', '#1A7BAA'
]


def getenv(name):
    v = os.getenv(name)
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def valid_email(addr):
    if not addr:
        return False
    s = str(addr).strip()
    if not s or any(c.isspace() for c in s):
        return False
    if s.count("@") != 1:
        return False
    local, _, domain = s.partition("@")
    if not local or not domain or "." not in domain:
        return False
    return True


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
    if not r.ok:
        print(f"  Graph error {r.status_code}: {r.text}")
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
    logo_path = "kream_logo.jpg"
    if os.path.exists(logo_path):
        img = Image.open(logo_path).convert('RGB')
        return np.array(img.resize((90, 90), Image.LANCZOS))
    return None


# ---------------- X (SocialPilot) integration ----------------

def derive_x_path(bsky_path):
    """X-Daten liegen in eigener Datei X_Follower_Log.xlsx im selben Ordner."""
    if "/" in bsky_path:
        folder = bsky_path.rsplit("/", 1)[0]
        return f"{folder}/X_Follower_Log.xlsx"
    return "X_Follower_Log.xlsx"


def load_x_data(token, sender, x_path):
    """Liest X_Log (Date, x_handle, Followers) + X_Accounts (loginId, x_handle, bluesky_handle, name).
    Returns (x_hist: {x_handle: {date: count}}, bsky2x: {bluesky_handle: [x_handle,...]}).
    Faellt bei jedem Fehler leise auf leer zurueck (Mail laeuft dann Bluesky-only).
    """
    x_hist = defaultdict(dict)
    bsky2x = defaultdict(list)
    try:
        meta = graph_get(token, f"{GRAPH}/users/{sender}/drive/root:/{x_path}")
        fid = meta["id"]
        sheets = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{fid}/workbook/worksheets")
        sid = {s["name"].lower(): s["id"] for s in sheets.get("value", [])}
        if "x_log" in sid:
            rng = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{fid}/workbook/worksheets/{sid['x_log']}/usedRange")
            for row in rng.get("values", [])[1:]:
                if len(row) < 3:
                    continue
                d = parse_excel_date(row[0])
                acc = str(row[1]).strip()
                try:
                    foll = int(float(str(row[2]).strip()))
                except Exception:
                    continue
                if d and acc:
                    x_hist[acc][d] = foll
        if "x_accounts" in sid:
            rng = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{fid}/workbook/worksheets/{sid['x_accounts']}/usedRange")
            for row in rng.get("values", [])[1:]:
                if len(row) < 3:
                    continue
                xh = str(row[1]).strip()
                bh = str(row[2]).strip()
                if xh and bh:
                    bsky2x[bh].append(xh)
    except Exception as e:
        print(f"  X data not loaded ({e}) -> Bluesky-only")
        return {}, {}
    return x_hist, bsky2x


def combined_x_series(x_hist, x_handles):
    """Mehrere X-Accounts eines Models zu einer Serie zusammenfuehren (carry-forward + Summe pro Datum)."""
    per = []
    all_dates = set()
    for xh in x_handles:
        pts = sorted(x_hist.get(xh, {}).items())
        if pts:
            per.append(pts)
            all_dates.update(d for d, _ in pts)
    if not per:
        return []
    out = []
    for d in sorted(all_dates):
        total = 0
        for pts in per:
            last = None
            for pd, pv in pts:
                if pd <= d:
                    last = pv
                else:
                    break
            if last is not None:
                total += last
        out.append((d, total))
    return out


def week_stats(day_data, today):
    sorted_days = sorted(day_data.items())
    dates = [d for d, _ in sorted_days]
    followers = [f for _, f in sorted_days]
    end = followers[-1]
    total_start = followers[0]
    total_growth = end - total_start
    total_pct = total_growth / total_start * 100 if total_start > 0 else 0
    cutoff = today - timedelta(days=7)
    week_start = None
    for d, f in sorted_days:
        if d <= cutoff:
            week_start = f
    if week_start is None:
        week_start = total_start
    week_growth = end - week_start
    week_pct = week_growth / week_start * 100 if week_start > 0 else 0
    return {
        "dates": dates, "followers": followers, "end": end,
        "week_start": week_start, "week_growth": week_growth, "week_pct": week_pct,
        "total_start": total_start, "total_growth": total_growth, "total_pct": total_pct,
    }


def fmt_signed(n):
    s = f'{abs(int(n)):,}'.replace(',', '.')
    return f'+{s}' if n >= 0 else f'-{s}'


def make_individual_chart(handle, name, day_data, today, logo_arr, tmpdir, x_series=None):
    st = week_stats(day_data, today)
    dates = st["dates"]
    followers = st["followers"]
    end = st["end"]
    growth = st["week_growth"]
    pct = st["week_pct"]

    fig, ax = plt.subplots(figsize=(11, 5.5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)

    ax.fill_between(dates, followers, alpha=0.06, color=ACCENT)
    ax.plot(dates, followers, color=ACCENT, linewidth=2.6, zorder=3, label='Bluesky')
    ax.scatter([dates[-1]], [followers[-1]], color=ACCENT, s=70, zorder=4)
    ax.annotate(f'{end:,}'.replace(',', '.'),
        xy=(dates[-1], followers[-1]), xytext=(10, 0),
        textcoords='offset points', fontsize=11, fontweight='bold', color=ACCENT, va='center')

    x_end = None
    if x_series:
        xd = [d for d, _ in x_series]
        xf = [f for _, f in x_series]
        x_end = xf[-1]
        ax.plot(xd, xf, color=X_COLOR, linewidth=2.6, zorder=3, label='X')
        ax.scatter([xd[-1]], [xf[-1]], color=X_COLOR, s=70, zorder=4)
        ax.annotate(f'{x_end:,}'.replace(',', '.'),
            xy=(xd[-1], xf[-1]), xytext=(10, 0),
            textcoords='offset points', fontsize=11, fontweight='bold', color=X_COLOR, va='center')

    growth_color = GREEN if growth >= 0 else RED
    fig.text(0.07, 0.90, name, color=CI_PRIMARY, fontsize=22, fontweight='bold', transform=fig.transFigure)
    fig.text(0.07, 0.835, f'@{handle}', color=TEXT2, fontsize=10, transform=fig.transFigure)
    head = f'Bluesky {end:,}'.replace(',', '.')
    if x_end is not None:
        head += f'   |   X {x_end:,}'.replace(',', '.')
    fig.text(0.55, 0.90, head, color=TEXT, fontsize=13, fontweight='bold', transform=fig.transFigure)
    sub = f'Bluesky {fmt_signed(growth)} this week ({fmt_signed(round(pct,1))}%)'
    if x_end is not None and len(x_series) < 2:
        sub += '   X tracking starts now'
    fig.text(0.55, 0.835, sub, color=growth_color, fontsize=10, transform=fig.transFigure)
    period = f'{dates[0].strftime("%d.%m.%Y")} to {dates[-1].strftime("%d.%m.%Y")}'
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
    if x_series:
        ax.legend(loc='upper left', framealpha=0.9, facecolor=BG, edgecolor=GRID, fontsize=10, labelcolor=TEXT)

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
            textcoords='offset points', fontsize=8.5, fontweight='bold', color=color, va='center')

    fig.text(0.05, 0.93, 'Bluesky Follower Overview', color=CI_PRIMARY, fontsize=18, fontweight='bold', transform=fig.transFigure)
    period = f'{min(all_dates).strftime("%d.%m.%Y")} to {max(all_dates).strftime("%d.%m.%Y")}'
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
    ax.legend(loc='upper left', framealpha=0.85, facecolor=BG, edgecolor=GRID, fontsize=9.5, labelcolor=TEXT)

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
    with open(attachment_path, 'rb') as f:
        content_bytes = base64.b64encode(f.read()).decode('utf-8')
    filename = os.path.basename(attachment_path)
    to_recipients = [{"emailAddress": {"address": to_email.strip()}}]
    cc_recipients = [
        {"emailAddress": {"address": r.strip()}}
        for r in cc_emails.split(",") if valid_email(r)
    ]
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": filename,
                "contentType": "image/png",
                "contentBytes": content_bytes
            }]
        },
        "saveToSentItems": True
    }
    graph_post(token, f"{GRAPH}/users/{sender}/sendMail", payload)


def build_model_html(name, handle, end_followers, week_growth, week_pct, period, x_end=None, x_handles_txt=""):
    fmt_end = f'{end_followers:,}'.replace(',', '.')
    fmt_growth = f'{fmt_signed(week_growth)} ({fmt_signed(round(week_pct,1))}%)'
    growth_color = '#1e7d1e' if week_growth >= 0 else '#cc0000'
    lbl = 'background:#005792; color:#fff; font-weight:600;'
    val = 'background:#eef3f7;'
    account_val = f'@{handle}'
    if x_handles_txt:
        account_val += f' &nbsp; {x_handles_txt}'
    x_row = ""
    total_row = ""
    if x_end is not None:
        x_row = f"""
          <tr>
            <td style="padding:10px 16px; {lbl}">X Followers</td>
            <td style="padding:10px 16px; {val}">{f'{x_end:,}'.replace(',', '.')}</td>
          </tr>"""
        total_row = f"""
          <tr>
            <td style="padding:10px 16px; {lbl}">Total</td>
            <td style="padding:10px 16px; {val} font-weight:700;">{f'{end_followers + x_end:,}'.replace(',', '.')} Followers</td>
          </tr>"""
    return f"""
    <!DOCTYPE html>
    <html><head><meta charset="utf-8"></head>
      <body style="font-family:'Source Sans 3','Source Sans Pro',Arial,sans-serif; color:#111; background:#f7f9fb; padding:24px;">
        <h2 style="margin:0 0 4px 0; color:#005792; font-family:'Exo 2',Arial,sans-serif;">Your Weekly Report</h2>
        <p style="color:#555; margin:0 0 20px 0;">{period}</p>
        <table style="border-collapse:collapse; width:480px;">
          <tr>
            <td style="padding:10px 16px; {lbl}">Account</td>
            <td style="padding:10px 16px; {val}">{account_val}</td>
          </tr>
          <tr>
            <td style="padding:10px 16px; {lbl}">Bluesky Followers</td>
            <td style="padding:10px 16px; {val}">{fmt_end} <span style="color:{growth_color}; font-weight:700;">{fmt_growth}</span></td>
          </tr>{x_row}{total_row}
        </table>
        <p style="margin:20px 0 4px 0; color:#555; font-size:13px;">Your growth chart is attached.</p>
        <p style="color:#555; font-size:12px; margin-top:24px;">krea-M</p>
      </body>
    </html>
    """


def build_manager_html(today, all_data, display_names):
    rows = ""
    sorted_accounts = sorted(all_data.items(), key=lambda x: max(x[1].values()), reverse=True)
    for handle, day_data in sorted_accounts:
        st = week_stats(day_data, today)
        name = display_names.get(handle, handle)
        wk_color = '#1e7d1e' if st["week_growth"] >= 0 else '#cc0000'
        tot_color = '#1e7d1e' if st["total_growth"] >= 0 else '#cc0000'
        fmt_end = f'{st["end"]:,}'.replace(',', '.')
        rows += f"""
        <tr>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;font-weight:600;">{name}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;color:#555;font-size:12px;">@{handle}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;">{fmt_end}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;color:{wk_color};font-weight:700;">{fmt_signed(st["week_growth"])}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;color:{wk_color};font-weight:700;">{fmt_signed(round(st["week_pct"],1))}%</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;color:{tot_color};">{fmt_signed(st["total_growth"])}</td>
          <td style="padding:10px 8px;border-bottom:1px solid #ddd;text-align:right;color:{tot_color};">{fmt_signed(round(st["total_pct"],1))}%</td>
        </tr>
        """
    return f"""
    <!DOCTYPE html>
    <html><head><meta charset="utf-8"></head>
      <body style="font-family:'Source Sans 3','Source Sans Pro',Arial,sans-serif; color:#111; background:#f7f9fb; padding:24px;">
        <h2 style="margin:0 0 4px 0; color:#005792;">Bluesky Weekly Overview - {today.strftime("%d.%m.%Y")}</h2>
        <p style="color:#555; margin:0 0 20px 0;">Weekly growth (last 7 days) and total growth since tracking start.</p>
        <table style="border-collapse:collapse; width:820px; max-width:100%;">
          <thead>
            <tr style="background:#005792;color:#fff;">
              <th style="text-align:left;padding:10px 8px;">Name</th>
              <th style="text-align:left;padding:10px 8px;">Handle</th>
              <th style="text-align:right;padding:10px 8px;">Followers</th>
              <th style="text-align:right;padding:10px 8px;">Week</th>
              <th style="text-align:right;padding:10px 8px;">Week %</th>
              <th style="text-align:right;padding:10px 8px;">Total</th>
              <th style="text-align:right;padding:10px 8px;">Total %</th>
            </tr>
          </thead>
          <tbody>{rows}</tbody>
        </table>
        <p style="color:#555; font-size:12px; margin-top:24px;">Overview chart attached.</p>
      </body>
    </html>
    """


def main():
    token = get_token()
    sender = getenv("SENDER_UPN")
    manager_emails = getenv("RECIPIENTS")
    path = getenv("ONEDRIVE_FILE_PATH")
    today = datetime.utcnow().date()

    file_meta = graph_get(token, f"{GRAPH}/users/{sender}/drive/root:/{path}")
    file_id = file_meta["id"]

    sheets = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/worksheets")
    models_sheet_id = None
    for s in sheets.get("value", []):
        if s["name"].lower() == "models":
            models_sheet_id = s["id"]
    if not models_sheet_id:
        raise RuntimeError("Worksheet 'Models' not found in Excel")

    models_range = graph_get(token, f"{GRAPH}/users/{sender}/drive/items/{file_id}/workbook/worksheets/{models_sheet_id}/usedRange")
    models = []
    display_names = {}
    for row in models_range.get("values", [])[1:]:
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

    # X-Daten (separat, defensiv)
    x_hist, bsky2x = load_x_data(token, sender, derive_x_path(path))

    logo_arr = load_logo()

    with tempfile.TemporaryDirectory() as tmpdir:
        for model in models:
            handle = model["handle"]
            name = model["name"]
            email = model["email"]
            if handle not in history or len(history[handle]) < 2:
                print(f"Skipping {handle} - not enough data")
                continue
            if not valid_email(email):
                print(f"Skipping {name} ({handle}) - invalid email: {email!r}")
                continue

            x_handles = bsky2x.get(handle, [])
            x_series = combined_x_series(x_hist, x_handles) if x_handles else []
            x_end = x_series[-1][1] if x_series else None
            x_handles_txt = "  ".join(f"@{h}" for h in x_handles) if x_handles else ""

            chart_path = make_individual_chart(handle, name, history[handle], today, logo_arr, tmpdir, x_series=x_series)

            st = week_stats(history[handle], today)
            dates = st["dates"]
            period = f'{dates[0].strftime("%d.%m.%Y")} to {dates[-1].strftime("%d.%m.%Y")}'
            html = build_model_html(name, handle, st["end"], st["week_growth"], st["week_pct"], period,
                                    x_end=x_end, x_handles_txt=x_handles_txt)
            subject = f"Your Bluesky Growth - {today.strftime('%d.%m.%Y')}"
            try:
                send_mail_with_attachment(token, sender, email, manager_emails, subject, html, chart_path)
                print(f"OK Mail sent to {name} ({email})")
            except Exception as e:
                print(f"FAIL send to {name} ({email}): {e}")
                continue

        all_data = {m["handle"]: history[m["handle"]] for m in models if m["handle"] in history and len(history[m["handle"]]) >= 2}
        overview_path = make_overview_chart(all_data, display_names, logo_arr, tmpdir)
        manager_html = build_manager_html(today, all_data, display_names)
        manager_list = [r.strip() for r in manager_emails.split(",") if valid_email(r)]
        if not manager_list:
            print("No valid manager email found - overview not sent")
        else:
            to_manager = manager_list[0]
            cc_managers = ",".join(manager_list[1:]) if len(manager_list) > 1 else ""
            try:
                send_mail_with_attachment(token, sender, to_manager, cc_managers,
                    f"Bluesky Weekly Overview - {today.strftime('%d.%m.%Y')}", manager_html, overview_path)
                print("OK Overview sent to managers")
            except Exception as e:
                print(f"FAIL overview: {e}")
    print("DONE")


if __name__ == "__main__":
    main()

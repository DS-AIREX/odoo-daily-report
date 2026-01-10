import os
import xmlrpc.client
import pandas as pd
from datetime import datetime
import pytz
from collections import defaultdict
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ========================
# LOAD SECRETS FROM GITHUB
# ========================
ODOO_URL = os.environ["ODOO_URL"]
ODOO_DB = os.environ["ODOO_DB"]
ODOO_USERNAME = os.environ["ODOO_USERNAME"]
ODOO_PASSWORD = os.environ["ODOO_PASSWORD"]

SENDER_EMAIL = os.environ["SENDER_EMAIL"]
RECEIVER_EMAIL = os.environ["RECEIVER_EMAIL"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]

# ========================
# ODOO CONNECTION
# ========================
common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD, {})
if not uid:
    raise Exception("❌ Odoo authentication failed")

print("✅ Connected to Odoo")

# ========================
# FETCH TODAY'S ACTIVITIES
# ========================
def get_daily_activities():
    ist = pytz.timezone("Asia/Kolkata")
    today = datetime.now(ist).date()

    start_utc = ist.localize(datetime.combine(today, datetime.min.time())).astimezone(pytz.utc)
    end_utc = ist.localize(datetime.combine(today, datetime.max.time())).astimezone(pytz.utc)

    activities = models.execute_kw(
        ODOO_DB, uid, ODOO_PASSWORD,
        "mail.activity", "search_read",
        [[
            ["res_model", "=", "crm.lead"],
            ["state", "=", "done"],
            ["write_date", ">=", start_utc.strftime("%Y-%m-%d %H:%M:%S")],
            ["write_date", "<=", end_utc.strftime("%Y-%m-%d %H:%M:%S")]
        ]],
        {"fields": ["user_id"], "limit": 2000}
    )

    counter = defaultdict(int)
    for act in activities:
        if act["user_id"]:
            counter[act["user_id"][1]] += 1

    return pd.DataFrame(
        [{"Sales Person": k, "Activities": v} for k, v in counter.items()]
    ).sort_values("Activities", ascending=False)

# ========================
# SEND EMAIL
# ========================
def send_email(df):
    html_table = df.to_html(index=False)

    today = datetime.now().strftime("%d %B %Y")
    total = df["Activities"].sum()

    html = f"""
    <h2>📊 Sales Activities Report – {today}</h2>
    <p><b>Total Activities:</b> {total}</p>
    {html_table}
    <br>
    <p>Auto-generated via GitHub Actions</p>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Sales Activities Report – {today}"
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(SENDER_EMAIL, EMAIL_PASSWORD)
        server.send_message(msg)

    print("✅ Email sent successfully")

# ========================
# RUN
# ========================
df = get_daily_activities()
if df.empty:
    print("⚠️ No activities found today")
else:
    send_email(df)

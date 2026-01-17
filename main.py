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
# FETCH TODAY'S COMPLETED ACTIVITIES (IMPROVED LOGIC)
# ========================
def get_daily_activities():
    ist = pytz.timezone("Asia/Kolkata")
    now_ist = datetime.now(ist)
    today_ist = now_ist.date()
    
    # Today at 12:00 AM (midnight)
    today_start_ist = ist.localize(datetime.combine(today_ist, datetime.min.time()))
    
    # Current time
    today_end_ist = now_ist
    
    # Convert to UTC for Odoo query
    start_utc = today_start_ist.astimezone(pytz.utc).strftime("%Y-%m-%d %H:%M:%S")
    end_utc = today_end_ist.astimezone(pytz.utc).strftime("%Y-%m-%d %H:%M:%S")
    
    print(f"📅 Today's Date: {today_ist}")
    print(f"⏰ Time Period: {today_start_ist.strftime('%I:%M %p')} to {today_end_ist.strftime('%I:%M %p')} IST")
    print(f"🌐 UTC Time Range: {start_utc} to {end_utc}")
    
    try:
        # Check if date_done field exists in mail.activity model
        fields = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'mail.activity', 'fields_get',
            [],
            {'attributes': ['string', 'type']}
        )
        
        date_done_exists = 'date_done' in fields
        print(f"📋 'date_done' field exists: {date_done_exists}")
        
        if date_done_exists:
            # Method 1: Use date_done field (most accurate)
            print("🔄 Using 'date_done' field for accurate completion time...")
            activities = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                "mail.activity", "search_read",
                [[
                    ["res_model", "=", "crm.lead"],
                    ["state", "=", "done"],
                    ["date_done", ">=", start_utc],
                    ["date_done", "<=", end_utc]
                ]],
                {
                    "fields": ["user_id", "date_done", "activity_type_id", "summary", "res_name"],
                    "limit": 2000,
                    "order": "date_done desc"
                }
            )
        else:
            # Method 2: Use write_date with filtering for actual completions
            print("🔄 Using 'write_date' field with filtering...")
            all_updated_activities = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                "mail.activity", "search_read",
                [[
                    ["res_model", "=", "crm.lead"],
                    ["state", "=", "done"],
                    ["write_date", ">=", start_utc],
                    ["write_date", "<=", end_utc]
                ]],
                {
                    "fields": ["user_id", "write_date", "create_date", "activity_type_id", 
                              "summary", "res_name", "state"],
                    "limit": 2000,
                    "order": "write_date desc"
                }
            )
            
            # Filter out activities that might have been created today but completed earlier
            activities = []
            for activity in all_updated_activities:
                create_date = activity.get("create_date")
                write_date = activity.get("write_date")
                
                if create_date and write_date:
                    create_dt = datetime.strptime(create_date, "%Y-%m-%d %H:%M:%S")
                    write_dt = datetime.strptime(write_date, "%Y-%m-%d %H:%M:%S")
                    
                    # Only include if created and marked done on same day
                    if create_dt.date() == write_dt.date() == today_ist:
                        activities.append(activity)
                else:
                    # If we can't determine, include it
                    activities.append(activity)
        
        print(f"✅ Total activities completed today: {len(activities)}")
        
        # Count activities per salesperson
        counter = defaultdict(int)
        for act in activities:
            if act["user_id"]:
                counter[act["user_id"][1]] += 1
        
        # Create DataFrame with same structure as old code
        df = pd.DataFrame(
            [{"Sales Person": k, "Activities": v} for k, v in counter.items()]
        ).sort_values("Activities", ascending=False)
        
        return df
        
    except Exception as e:
        print(f"❌ Error fetching activities: {str(e)}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

# ========================
# SEND EMAIL (SAME AS BEFORE)
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
# RUN (SAME STRUCTURE AS BEFORE)
# ========================
if __name__ == "__main__":
    df = get_daily_activities()
    if df.empty:
        print("⚠️ No activities found today")
    else:
        send_email(df)

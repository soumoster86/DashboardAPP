import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText

# -----------------------------
# CONFIG
# -----------------------------
st.set_page_config(page_title="Upgrade Dashboard", layout="wide")

st.title("📊 Upgrade Dashboard")
st.caption("Track progress • Identify critical users • Notify managers")

# -----------------------------
# HIGHLIGHT FUNCTION
# -----------------------------
def highlight_status(val):
    colors = {
        "Not Booked": "background-color:#ff4d4d;color:white",
        "Updated": "background-color:#28a745;color:white",
        "Booked": "background-color:#ffa500",
        "Unreachable": "background-color:#6c757d;color:white",
        "Left Org": "background-color:#000;color:white"
    }
    return colors.get(val, "")

# -----------------------------
# LOAD DATA (CACHED)
# -----------------------------
@st.cache_data
def load_data(file):
    return pd.read_excel(file, engine="openpyxl")

# -----------------------------
# EMAIL FUNCTION
# -----------------------------
def send_email(to_email, subject, html_body):
    try:
        sender_email = st.secrets["email"]
        password = st.secrets["password"]
        smtp_server = st.secrets.get("smtp_server", "smtp.office365.com")
        smtp_port = st.secrets.get("smtp_port", 587)

        msg = MIMEText(html_body, "html")
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = to_email

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(msg)

        return True

    except Exception as e:
        st.error(f"❌ Email failed: {e}")
        return False

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:

    try:
        df = load_data(uploaded_file)
    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        st.stop()

    # -----------------------------
    # VALIDATION
    # -----------------------------
    df.columns = df.columns.str.strip()

    required_cols = ["Name", "PC name", "Status", "Manager", "Manager Email"]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"⚠️ Missing columns: {', '.join(missing_cols)}")
        st.stop()

    # -----------------------------
    # CLEANING
    # -----------------------------
    df["Status"] = (
        df["Status"]
        .fillna("Not Booked")
        .astype(str)
        .str.strip()
        .replace("", "Not Booked")
        .str.title()
    )

    df["Manager"] = df["Manager"].fillna("Unknown").astype(str).str.strip()
    df["Manager Email"] = df["Manager Email"].fillna("").astype(str).str.strip()

    df["Status"] = df["Status"].replace({
        "Nan": "Not Booked",
        "None": "Not Booked",
        "Notbooked": "Not Booked",
        "Left": "Left Org",
        "Leftorg": "Left Org",
        "No Response": "Unreachable"
    })

    valid_status = ["Updated", "Booked", "Not Booked", "Left Org", "Unreachable"]

    df["Status"] = df["Status"].apply(lambda x: x if x in valid_status else "Not Booked")

    total_users = len(df)

    # -----------------------------
    # METRICS
    # -----------------------------
    status_counts = df["Status"].value_counts()

    st.subheader("📊 Status Summary")

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("✅ Updated", status_counts.get("Updated", 0))
    col2.metric("📅 Booked", status_counts.get("Booked", 0))
    col3.metric("❌ Not Booked", status_counts.get("Not Booked", 0))
    col4.metric("👋 Left Org", status_counts.get("Left Org", 0))
    col5.metric("📵 Unreachable", status_counts.get("Unreachable", 0))

    # -----------------------------
    # PROGRESS
    # -----------------------------
    if total_users > 0:
        completion = (status_counts.get("Updated", 0) / total_users) * 100
        st.progress(completion / 100)
        st.write(f"### 🚀 Completion: {completion:.2f}%")

    # -----------------------------
    # CRITICAL USERS
    # -----------------------------
    st.subheader("🚨 Critical Users")

    critical_df = df[df["Status"].isin(["Not Booked", "Unreachable"])]

    styled_critical = critical_df.style.map(highlight_status, subset=["Status"])
    st.dataframe(styled_critical, use_container_width=True)

    # -----------------------------
    # MANAGER SUMMARY
    # -----------------------------
    st.subheader("👨‍💼 Manager Breakdown")

    manager_summary = (
        critical_df.groupby(["Manager", "Manager Email"])
        .size()
        .reset_index(name="Critical Count")
        .sort_values(by="Critical Count", ascending=False)
    )

    st.dataframe(manager_summary, use_container_width=True)

    # -----------------------------
    # FILTERS
    # -----------------------------
    st.subheader("🔍 Filters")

    col1, col2 = st.columns(2)

    managers = sorted(df["Manager"].unique())
    statuses = sorted(df["Status"].unique())

    selected_manager = col1.selectbox("Manager", ["All"] + managers)
    selected_status = col2.selectbox("Status", ["All"] + statuses)

    filtered_df = df.copy()

    if selected_manager != "All":
        filtered_df = filtered_df[filtered_df["Manager"] == selected_manager]

    if selected_status != "All":
        filtered_df = filtered_df[filtered_df["Status"] == selected_status]

    styled_filtered = filtered_df.style.map(highlight_status, subset=["Status"])
    st.dataframe(styled_filtered, use_container_width=True)

    # -----------------------------
    # EMAIL PREVIEW
    # -----------------------------
    st.subheader("📧 Email Preview")

    if st.checkbox("Show Email Preview"):
        for (manager, email), group in critical_df.groupby(["Manager", "Manager Email"]):
            st.write(f"### {manager} ({email})")
            st.dataframe(group[["Name", "PC name", "Status"]])

    # -----------------------------
    # SEND EMAILS
    # -----------------------------
    st.subheader("📨 Send Notifications")

    if st.button("Send Emails to Managers"):

        if critical_df.empty:
            st.info("No critical users. No emails sent.")
        else:
            success = 0

            for (manager, email), group in critical_df.groupby(["Manager", "Manager Email"]):

                if not email:
                    st.warning(f"⚠️ Missing email for {manager}")
                    continue

                html_table = group[["Name", "PC name", "Status"]].to_html(index=False)

                body = f"""
                <h3>Hi {manager},</h3>
                <p>The following users require your attention:</p>
                {html_table}
                <br>
                <p>Please take necessary action.</p>
                <p>Regards,<br>IT Team</p>
                """

                if send_email(email, "🚨 Pending Windows Upgrade Users", body):
                    success += 1

            st.success(f"✅ Emails sent to {success} managers")

    # -----------------------------
    # VISUALIZATION
    # -----------------------------
    st.subheader("📊 Visualization")
    st.bar_chart(status_counts)

    st.subheader("🔥 Top Managers with Issues")
    st.bar_chart(critical_df["Manager"].value_counts())

    # -----------------------------
    # DOWNLOAD
    # -----------------------------
    st.subheader("⬇️ Download Summary")

    csv = status_counts.to_csv().encode("utf-8")

    st.download_button(
        label="Download CSV",
        data=csv,
        file_name="status_summary.csv",
        mime="text/csv"
    )

    # -----------------------------
    # FOOTER
    # -----------------------------
    st.caption(f"Last refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

else:
    st.info("👆 Upload Excel file to begin")
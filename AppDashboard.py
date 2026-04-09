import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
import matplotlib.pyplot as plt

# -----------------------------
# CONFIG
# -----------------------------
st.set_page_config(page_title="Upgrade Dashboard", layout="wide")

st.title("📊 Windows 11 Upgrade Dashboard")
st.caption("Actionable insights • Manager accountability • Automated notifications")

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
# LOAD DATA
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

        msg = MIMEText(html_body, "html")
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = to_email

        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(msg)

        return True

    except Exception as e:
        st.error(f"Email failed: {e}")
        return False

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:

    df = load_data(uploaded_file)
    df.columns = df.columns.str.strip()

    # -----------------------------
    # VALIDATION
    # -----------------------------
    required_cols = ["Name", "PC name", "Status", "Manager", "Manager Email"]

    if any(col not in df.columns for col in required_cols):
        st.error("Missing required columns!")
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
        "Left": "Left Org",
        "No Response": "Unreachable"
    })

    valid_status = ["Updated", "Booked", "Not Booked", "Left Org", "Unreachable"]
    df["Status"] = df["Status"].apply(lambda x: x if x in valid_status else "Not Booked")

    total_users = len(df)
    status_counts = df["Status"].value_counts()

    # -----------------------------
    # METRICS
    # -----------------------------
    st.subheader("📊 Key Metrics")

    col1, col2, col3 = st.columns(3)

    completion_pct = (status_counts.get("Updated", 0) / total_users) * 100
    pending_pct = 100 - completion_pct

    col1.metric("🚀 Completion %", f"{completion_pct:.1f}%")
    col2.metric("⚠️ Pending %", f"{pending_pct:.1f}%")
    col3.metric("👥 Total Users", total_users)

    # -----------------------------
    # STATUS DISTRIBUTION
    # -----------------------------
    st.subheader("📊 Status Distribution")
    st.bar_chart(status_counts.sort_values(ascending=False))

    # -----------------------------
    # PIE CHART
    # -----------------------------
    st.subheader("🥧 Status Breakdown")

    fig, ax = plt.subplots()
    ax.pie(status_counts, labels=status_counts.index, autopct='%1.1f%%')
    ax.axis('equal')
    st.pyplot(fig)

    # -----------------------------
    # CRITICAL USERS
    # -----------------------------
    st.subheader("🚨 Critical Users")

    critical_df = df[df["Status"].isin(["Not Booked", "Unreachable"])]

    styled = critical_df.style.map(highlight_status, subset=["Status"])
    st.dataframe(styled, use_container_width=True)

    # -----------------------------
    # MANAGER INSIGHTS
    # -----------------------------
    st.subheader("👨‍💼 Manager-wise Issues")

    manager_issues = critical_df["Manager"].value_counts()
    st.bar_chart(manager_issues)

    # -----------------------------
    # MANAGER VS STATUS
    # -----------------------------
    st.subheader("📊 Manager vs Status Matrix")

    manager_status = pd.crosstab(df["Manager"], df["Status"])
    st.bar_chart(manager_status)

    # -----------------------------
    # TOP 5 MANAGERS
    # -----------------------------
    st.subheader("🔥 Top 5 Managers Requiring Attention")

    top5 = manager_issues.head(5)
    st.bar_chart(top5)

    # -----------------------------
    # FILTERS
    # -----------------------------
    st.subheader("🔍 Filters")

    col1, col2 = st.columns(2)

    selected_manager = col1.selectbox("Manager", ["All"] + sorted(df["Manager"].unique()))
    selected_status = col2.selectbox("Status", ["All"] + sorted(df["Status"].unique()))

    filtered_df = df.copy()

    if selected_manager != "All":
        filtered_df = filtered_df[filtered_df["Manager"] == selected_manager]

    if selected_status != "All":
        filtered_df = filtered_df[filtered_df["Status"] == selected_status]

    styled_filtered = filtered_df.style.map(highlight_status, subset=["Status"])
    st.dataframe(styled_filtered, use_container_width=True)

    # -----------------------------
    # EMAIL SECTION
    # -----------------------------
    st.subheader("📧 Notify Managers")

    if st.button("Send Emails"):

        success = 0

        for (manager, email), group in critical_df.groupby(["Manager", "Manager Email"]):

            if not email:
                continue

            html_table = group[["Name", "PC name", "Status"]].to_html(index=False)

            body = f"""
            <h3>Hi {manager},</h3>
            <p>Below users need action:</p>
            {html_table}
            <br>
            <p>Please take necessary action.</p>
            """

            if send_email(email, "Pending Upgrade Users", body):
                success += 1

        st.success(f"Emails sent: {success}")

    # -----------------------------
    # DOWNLOAD
    # -----------------------------
    st.subheader("⬇️ Download Summary")

    csv = status_counts.to_csv().encode("utf-8")

    st.download_button("Download CSV", csv, "summary.csv", "text/csv")

    # -----------------------------
    # FOOTER
    # -----------------------------
    st.caption(f"Last refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

else:
    st.info("Upload Excel file to begin")
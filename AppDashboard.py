import streamlit as st
import pandas as pd

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(page_title="Upgrade Dashboard", layout="wide")

st.title("📊 Upgrade Status Dashboard")

# -----------------------------
# Highlight Function
# -----------------------------
def highlight_status(val):
    if val == "Not Booked":
        return "background-color: red; color: white"
    elif val == "Updated":
        return "background-color: green; color: white"
    elif val == "Booked":
        return "background-color: orange"
    elif val == "Unreachable":
        return "background-color: gray; color: white"
    elif val == "Left Org":
        return "background-color: black; color: white"
    return ""

# -----------------------------
# File Upload
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:

    # Read Excel safely
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # Clean column names
    df.columns = df.columns.str.strip()

    # Validate column
    if "Status" not in df.columns:
        st.error("⚠️ 'Status' column not found!")
        st.stop()

    # -----------------------------
    # Data Cleaning
    # -----------------------------
    df["Status"] = (
        df["Status"]
        .fillna("Not Booked")
        .astype(str)
        .str.strip()
        .replace("", "Not Booked")
        .str.title()
    )

    df["Status"] = df["Status"].replace({
        "Nan": "Not Booked",
        "None": "Not Booked",
        "Notbooked": "Not Booked",
        "Not Booked ": "Not Booked",
        "Left": "Left Org",
        "Leftorg": "Left Org",
        "No Response": "Unreachable"
    })

    valid_status = ["Updated", "Booked", "Not Booked", "Left Org", "Unreachable"]

    df["Status"] = df["Status"].apply(
        lambda x: x if x in valid_status else "Not Booked"
    )

    # -----------------------------
    # Status Counts
    # -----------------------------
    status_counts = df["Status"].value_counts()

    updated = status_counts.get("Updated", 0)
    booked = status_counts.get("Booked", 0)
    not_booked = status_counts.get("Not Booked", 0)
    left_org = status_counts.get("Left Org", 0)
    unreachable = status_counts.get("Unreachable", 0)

    total_users = len(df)

    # -----------------------------
    # Metrics
    # -----------------------------
    st.subheader("📊 Status Summary")

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("✅ Updated", updated)
    col2.metric("📅 Booked", booked)
    col3.metric("❌ Not Booked", not_booked)
    col4.metric("👋 Left Org", left_org)
    col5.metric("📵 Unreachable", unreachable)

    # -----------------------------
    # Critical Users
    # -----------------------------
    critical_df = df[df["Status"].isin(["Not Booked", "Unreachable"])]

    st.subheader("🚨 Critical Users (Action Required)")
    styled_critical = critical_df.style.map(highlight_status, subset=["Status"])
    st.dataframe(styled_critical)

    # -----------------------------
    # Status Percentage
    # -----------------------------
    st.subheader("📊 Status Percentage")
    percent_df = (status_counts / total_users * 100).round(2)
    st.write(percent_df)

    # -----------------------------
    # Progress Bar
    # -----------------------------
    if total_users > 0:
        completion = (updated / total_users) * 100
        st.progress(completion / 100)
        st.write(f"### 🚀 Upgrade Completion: {completion:.2f}%")

    # -----------------------------
    # Raw Data (Styled)
    # -----------------------------
    st.subheader("📄 Raw Data")
    styled_df = df.style.map(highlight_status, subset=["Status"])
    st.dataframe(styled_df)

    # -----------------------------
    # Filter Section
    # -----------------------------
    st.subheader("🔍 Filter Data")

    status_list = sorted(set(df["Status"]))

    selected_status = st.selectbox(
        "Filter by Status",
        ["All"] + list(status_list)
    )

    if selected_status != "All":
        filtered_df = df[df["Status"] == selected_status]
    else:
        filtered_df = df

    styled_filtered = filtered_df.style.map(highlight_status, subset=["Status"])
    st.dataframe(styled_filtered)

    # -----------------------------
    # Detailed Counts
    # -----------------------------
    st.subheader("📈 Detailed Counts")
    st.write(status_counts)

    # -----------------------------
    # Chart
    # -----------------------------
    st.subheader("📊 Visualization")
    st.line_chart(status_counts)

    # -----------------------------
    # Download Summary
    # -----------------------------
    st.subheader("⬇️ Download Summary")

    csv = status_counts.to_csv().encode("utf-8")

    st.download_button(
        label="Download Status Summary",
        data=csv,
        file_name="status_summary.csv",
        mime="text/csv"
    )

else:
    st.info("👆 Please upload an Excel file to begin.")
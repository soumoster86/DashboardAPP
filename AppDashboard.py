import streamlit as st
import pandas as pd

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(page_title="Windows 11 Upgrade Dashboard", layout="wide")

st.title("📊 Windows 11 Upgrade Status Dashboard")

# -----------------------------
# File Upload
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Clean column names
    df.columns = df.columns.str.strip()

    # Validate column
    if "Status" not in df.columns:
        st.error("⚠️ 'Status' column not found!")
        st.stop()

    # -----------------------------
    # Data Cleaning (BULLETPROOF)
    # -----------------------------
    df["Status"] = (
        df["Status"]
        .fillna("Not Booked")        # Handle NaN
        .astype(str)                # Convert everything to string
        .str.strip()                # Remove spaces
        .replace("", "Not Booked")  # Empty → Not Booked
        .str.title()                # Normalize case
    )

    # Fix weird values
    df["Status"] = df["Status"].replace({
        "Nan": "Not Booked",
        "None": "Not Booked",
        "Notbooked": "Not Booked",
        "Not Booked ": "Not Booked",
        "Left": "Left Org",
        "Leftorg": "Left Org",
        "No Response": "Unreachable"
    })

    # Allowed statuses (strict control)
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
    # Progress Bar
    # -----------------------------
    if total_users > 0:
        completion = (updated / total_users) * 100
        st.progress(completion / 100)
        st.write(f"### 🚀 Upgrade Completion: {completion:.2f}%")

    # -----------------------------
    # Raw Data
    # -----------------------------
    st.subheader("📄 Raw Data")
    st.dataframe(df)

    # -----------------------------
    # Filter Section (FIXED ERROR)
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

    st.dataframe(filtered_df)

    # -----------------------------
    # Detailed Counts
    # -----------------------------
    st.subheader("📈 Detailed Counts")
    st.write(status_counts)

    # -----------------------------
    # Chart
    # -----------------------------
    st.subheader("📊 Visualization")
    st.bar_chart(status_counts)

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
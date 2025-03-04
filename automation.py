import streamlit as st
import pandas as pd
import plotly.express as px

# Load Excel Data
FILE_PATH = "C:/Users/asus/OneDrive - NIIT Limited/Automation/Destination_data.xlsx"
df = pd.read_excel(FILE_PATH, sheet_name="Formatted Data", engine="openpyxl")

# Data Cleaning
df.columns = df.columns.str.strip()
df["No.of Hours"] = pd.to_numeric(df["No.of Hours"], errors='coerce').fillna(0)
df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
df["Month"] = df["Date"].dt.month_name()
df["Day"] = df["Date"].dt.day_name()

# Define month order
month_order = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Streamlit UI
st.set_page_config(page_title="📊 Session Dashboard", layout="wide")

# Authentication
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login():
    st.title("🔐 Login Page")
    username = st.text_input("👤 Enter Username:")
    password = st.text_input("🔑 Enter Password:", type="password")
    if st.button("🔓 Login"):
        if username.lower() == "arya" and password == "12345":
            st.session_state.authenticated = True
            st.success("✅ Login Successful!")
            st.rerun()
        else:
            st.error("❌ Incorrect Username or Password!")

if not st.session_state.authenticated:
    login()
    st.stop()

# Sidebar Navigation
st.sidebar.title("📌 Navigation")
page = st.sidebar.radio("Go to", ["📊 Dashboard", "📄 Data Viewer", "⚠️ Session Clashes"])

# 📊 Dashboard Page
if page == "📊 Dashboard":
    st.title("📊 Session Dashboard")

    # Dropdown Filters
    col1, col2, col3 = st.columns(3)
    with col1:
        selected_month = st.selectbox(
            "📅 Select Month",
            ["All"] + [m for m in month_order if m in df["Month"].unique()]
        )
    
    with col2:
        selected_mentor = st.selectbox("👨‍🏫 Select Mentor", ["All"] + sorted(df["Mentor / Faculty"].dropna().unique()))
    
    with col3:
        selected_program = st.selectbox("📚 Select Program", ["All"] + sorted(df["Program Name"].dropna().unique()))
    
    # Filter Data
    filtered_df = df.copy()
    if selected_month != "All":
        filtered_df = filtered_df[filtered_df["Month"] == selected_month]
    if selected_mentor != "All":
        filtered_df = filtered_df[filtered_df["Mentor / Faculty"] == selected_mentor]
    if selected_program != "All":
        filtered_df = filtered_df[filtered_df["Program Name"] == selected_program]

    # Display Metrics
    total_sessions = len(filtered_df)
    total_hours = filtered_df["No.of Hours"].sum()
    st.markdown(f"📅 **Total Sessions in {selected_month}:** {total_sessions}")
    st.markdown(f"⏳ **Total Hours:** {total_hours}")

    # Display Chart
    if not filtered_df.empty:
        mentor_sessions = filtered_df["Mentor / Faculty"].value_counts().reset_index()
        mentor_sessions.columns = ["Mentor / Faculty", "Sessions"]
        fig = px.bar(mentor_sessions, x="Mentor / Faculty", y="Sessions", title="Sessions per Mentor", color="Sessions", text_auto=True)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("⚠️ No data available for the selected filters.")

# 📄 Data Viewer Page
elif page == "📄 Data Viewer":
    st.title("📄 View & Download Data")

    # Dropdown Filters
    col1, col2 = st.columns(2)
    with col1:
        selected_month = st.selectbox(
            "📅 Filter by Month",
            ["All"] + [m for m in month_order if m in df["Month"].unique()],
            key="month_view"
        )
    
    with col2:
        selected_mentor = st.selectbox(
            "👨‍🏫 Filter by Mentor",
            ["All"] + sorted(df["Mentor / Faculty"].dropna().unique()),
            key="mentor_view"
        )
    
    col3, col4 = st.columns(2)
    with col3:
        selected_day = st.selectbox(
            "📆 Filter by Day",
            ["All"] + sorted(df["Day"].dropna().unique()),
            key="day_view"
        )
    
    with col4:
        selected_client = st.selectbox(
            "🏢 Filter by Client",
            ["All"] + sorted(df["Client"].dropna().unique()) if "Client" in df.columns else ["N/A"],
            key="client_view"
        )

    # Apply Filters
    filtered_data = df.copy()
    if selected_month != "All":
        filtered_data = filtered_data[filtered_data["Month"] == selected_month]
    if selected_mentor != "All":
        filtered_data = filtered_data[filtered_data["Mentor / Faculty"] == selected_mentor]
    if selected_day != "All":
        filtered_data = filtered_data[filtered_data["Day"] == selected_day]
    if "Client" in df.columns and selected_client != "All":
        filtered_data = filtered_data[filtered_data["Client"] == selected_client]

    # Display Filtered Data
    st.dataframe(filtered_data, use_container_width=True)

    # Download Filtered Data
    csv_data = filtered_data.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Filtered Data",
        data=csv_data,
        file_name="filtered_data.csv",
        mime="text/csv"
    )

# ⚠️ Session Clashes Page
elif page == "⚠️ Session Clashes":
    st.title("⚠️ Session Clashes Detection")
    clashes = df[df.duplicated(subset=["Date", "Session Time", "Mentor / Faculty"], keep=False)]
    if not clashes.empty:
        st.warning("⚠️ The following session clashes have been detected:")
        st.dataframe(clashes, use_container_width=True)
    else:
        st.success("✅ No session clashes detected!")

# Logout Button
st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout"):
    st.session_state.authenticated = False
    st.rerun()

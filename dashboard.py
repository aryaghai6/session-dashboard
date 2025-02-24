import streamlit as st
import pandas as pd

# File path for the destination data
DATA_FILE = "C:/Users/asus/OneDrive - NIIT Limited/Automation/Destination_data.xlsx"
SHEET_NAME = "Formatted Data"

# Function to load data
def load_data():
    try:
        df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()

# Function to check session clashes
def check_clashes(df):
    clashes = df[df.duplicated(subset=["Date", "Session Time"], keep=False)]
    return clashes

# Function to find empty slots
def find_empty_slots(df):
    df_sorted = df.sort_values(by=["Date", "Session Time"])
    return df_sorted  # Placeholder for empty slot detection logic

# Streamlit UI
st.title("Session Dashboard")

# Simple Login
username = st.text_input("Enter your name to continue:")
if st.button("Login"):
    if username.lower() == "arya":
        st.success("Welcome, Arya!")
        
        # Load and display data
        df = load_data()
        if not df.empty:
            st.subheader("Session Details")
            st.dataframe(df)
            
            # Display KPIs
            st.subheader("Key Performance Indicators")
            st.write(f"Total Sessions: {len(df)}")
            st.write(f"Total Hours: {df['No. of Hours'].sum()}")
            
            # Check for session clashes
            st.subheader("Session Clashes")
            clashes = check_clashes(df)
            if not clashes.empty:
                st.write("⚠️ Overlapping sessions detected!")
                st.dataframe(clashes)
            else:
                st.write("✅ No session clashes detected.")
            
            # Check for empty slots
            st.subheader("Empty Slots")
            empty_slots = find_empty_slots(df)
            st.dataframe(empty_slots)
    else:
        st.error("Access denied. Only Arya can log in.")
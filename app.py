import streamlit as st
import pandas as pd
import msal
import requests

# 🔐 Load secrets from Streamlit Secrets Manager
CLIENT_ID = st.secrets["graph"]["client_id"]
CLIENT_SECRET = st.secrets["graph"]["client_secret"]
TENANT_ID = st.secrets["graph"]["tenant_id"]
DRIVE_ID = st.secrets["graph"]["drive_id"]
ITEM_ID = st.secrets["graph"]["item_id"]

# 📌 API base components
BASE_URL = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook"

# 🔑 Get access token
def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token["access_token"]

# 📥 Read data from data1 sheet
def get_data1(token):
    url = f"{BASE_URL}/worksheets/data1/usedRange"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers).json()

    if "values" in resp:
        rows = resp["values"]
        return pd.DataFrame(rows[1:], columns=rows[0])
    else:
        st.error("❌ Could not retrieve 'values' from data1. Check table format.")
        return pd.DataFrame()

# ➕ Append row to data2 table
def append_to_data2(token, district, name, literacy):
    url = f"{BASE_URL}/worksheets/data2/tables/Table1/rows/add"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = {
        "values": [[district, name, literacy]]
    }
    resp = requests.post(url, headers=headers, json=body)
    return resp.status_code == 201

# 🧠 Streamlit UI
st.title("📋 Literacy Data Entry Form")

token = get_token()
df = get_data1(token)

if df.empty:
    st.stop()

# 🎯 Dynamic Dropdowns
districts = df["District"].dropna().unique()
selected_district = st.selectbox("Select District", sorted(districts))

names = df[df["District"] == selected_district]["Name"].dropna().unique()
selected_name = st.selectbox("Select Name", sorted(names))

literacy_status = st.text_input("Enter Literacy Status")

# 🔘 Submission
if selected_district and selected_name and literacy_status:
    if st.button("Submit"):
        success = append_to_data2(token, selected_district, selected_name, literacy_status)
        if success:
            st.success("✅ Entry submitted successfully to Excel Online!")
        else:
            st.error("❌ Submission failed. Please check API permissions or table setup.")

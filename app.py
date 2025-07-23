import streamlit as st
import pandas as pd
import msal
import requests

# 🔐 Load credentials from Streamlit Secrets Manager
CLIENT_ID     = st.secrets["graph"]["client_id"]
CLIENT_SECRET = st.secrets["graph"]["client_secret"]
TENANT_ID     = st.secrets["graph"]["tenant_id"]
DRIVE_ID      = st.secrets["graph"]["drive_id"]
ITEM_ID       = st.secrets["graph"]["item_id"]

BASE_URL = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook"

# 🔑 Acquire Microsoft Graph access token
def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token_response["access_token"]

# 📥 Read rows from Data1Table
def get_data1(token):
    url = f"{BASE_URL}/tables/Data1Table/rows"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers).json()

    if "value" in resp:
        rows = [row["values"][0] for row in resp["value"]]
        columns = ["District", "Name", "Gender"]
        return pd.DataFrame(rows, columns=columns)
    else:
        st.error("❌ Could not retrieve rows from Data1Table. Check table name and structure.")
        return pd.DataFrame()

# ➕ Append a new row to Table1 in data2 worksheet
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

# 🧠 Streamlit Interface
st.title("📋 Literacy Data Entry")

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

# 🔘 Submission Button
if selected_district and selected_name and literacy_status:
    if st.button("Submit"):
        success = append_to_data2(token, selected_district, selected_name, literacy_status)
        if success:
            st.success("✅ Entry successfully submitted to Excel Online!")
        else:
            st.error("❌ Submission failed. Check table permissions or formatting.")

import streamlit as st
import pandas as pd
import msal
import requests

# 🔐 Load secrets from Streamlit's Secrets Manager
CLIENT_ID = st.secrets["graph"]["client_id"]
CLIENT_SECRET = st.secrets["graph"]["client_secret"]
TENANT_ID = st.secrets["graph"]["tenant_id"]
EXCEL_FILE_ID = st.secrets["graph"]["excel_file_id"]

# 🔑 Get access token from Microsoft Graph
def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token["access_token"]

# 📥 Read data1 sheet
def get_data1(token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{EXCEL_FILE_ID}/workbook/worksheets/data1/usedRange"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers).json()
    
    st.write("📦 Graph API Response:")
    st.json(resp)  # Displays the full JSON response

    if "values" in resp:
        rows = resp["values"]
        df = pd.DataFrame(rows[1:], columns=rows[0])
        return df
    else:
        st.error("❌ 'values' not found in API response. Check sheet name or format.")
        return pd.DataFrame()  # Return empty to prevent app crash


# ➕ Append to data2 sheet
def append_data2(token, district, name, literacy):
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{EXCEL_FILE_ID}/workbook/worksheets/data2/tables/Table1/rows/add"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = {
        "values": [[district, name, literacy]]
    }
    response = requests.post(url, headers=headers, json=body)
    return response.status_code == 201

# 🧠 Streamlit UI
st.title("📋 Literacy Data Entry Dashboard")

token = get_token()
df = get_data1(token)

district = st.selectbox("Select District", df["District"].unique())
names = df[df["District"] == district]["Name"].unique()
name = st.selectbox("Select Name", names)
literacy = st.text_input("Enter Literacy Status")

if district and name and literacy:
    if st.button("Submit"):
        success = append_data2(token, district, name, literacy)
        if success:
            st.success("✅ Data submitted to Excel Online!")
        else:
            st.error("❌ Submission failed. Please try again.")

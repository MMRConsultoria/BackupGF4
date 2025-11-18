import json, streamlit as st
from oauth2client.service_account import ServiceAccountCredentials

secret = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
credentials_dict = json.loads(secret) if isinstance(secret, str) else dict(secret)
st.write("Service account:", credentials_dict.get("client_email"))

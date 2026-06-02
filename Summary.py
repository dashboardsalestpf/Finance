import streamlit as st
st.set_page_config(layout="wide")

import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import altair as alt
import json
import datetime
import time
import io
from google.oauth2.service_account import Credentials
import gspread

@st.cache_resource
def connect_gsheet(x):
    creds_dict = st.secrets["projectfinance"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1TbrJybGgTjJowu2sfFw4s3NFcAD580y4O8MvEEZtWuw").worksheet(x)
    return sheet

# ----- Load Google Sheet into DataFrame -----
def load_data(x):
    sheet = connect_gsheet(x)
    values = sheet.get_all_values()
    headers = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)
    return df

@st.cache_data
def loading_data():
    df = load_data("Sheet1")
    df2 = load_data("m_branch")
    df3 = load_data("m_source")
    df4 = load_data("m_vendor")
    return df, df2, df3, df4

df, df2, df3, df4 = loading_data()

st.title("Data from Google Sheet")

editable_cols = ["Invoice", "FP", "Surat Jalan", "Refund", "Notes"]

edited_df = st.data_editor(
    df,
    use_container_width=True,
    column_config={
        "Invoice": st.column_config.SelectboxColumn(
            "Invoice",
            options=["Not Done", "Done"],
        ),
        "FP": st.column_config.SelectboxColumn(
            "FP",
            options=["Not Done", "Done"],
        ),
        "Surat Jalan": st.column_config.SelectboxColumn(
            "Surat Jalan",
            options=["Not Done", "Done"],
        ),
        "Refund": st.column_config.SelectboxColumn(
            "Refund",
            options=["No Refund", "On Progress", "Done"],
        ),
        "Notes": st.column_config.TextColumn(
            "Notes"
        ),
    },
    disabled=[col for col in df.columns if col not in editable_cols]
)

if st.button("Save Changes"):
    sheet = connect_gsheet("Sheet1")

    sheet.update(
        [edited_df.columns.values.tolist()] +
        edited_df.fillna("").values.tolist()
    )
    st.success("Data updated to Google Sheet!")
    st.cache_data.clear()
    st.cache_resource.clear()
    st.switch_page("Summary.py")
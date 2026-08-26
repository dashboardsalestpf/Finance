import streamlit as st
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import altair as alt
import json
import datetime
import time
import io
from google.oauth2.service_account import Credentials
import gspread

today = datetime.datetime.today()
yyyymm = today.strftime("%Y%m")


@st.cache_resource
def connect_gsheet(x):
    creds_dict = st.secrets["projectfinance"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1TbrJybGgTjJowu2sfFw4s3NFcAD580y4O8MvEEZtWuw").worksheet(x)
    return sheet

EXPECTED_COLUMNS = [
    "Branch",
    "Source",
    "Tanggal Terima",
    "Vendor",
    "No Invoice",
    "Keterangan",
    "Nominal",
    "Payment Date",
    "Payment Code",
    "Invoice",
    "FP",
    "Surat Jalan",
    "Refund",
    "Notes",
    "Nomor Ticket"
]

def load_data(x):
    sheet = connect_gsheet(x)
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    # force schema even if sheet is empty
    if df.empty:
        df = pd.DataFrame(columns=EXPECTED_COLUMNS)
    return df

@st.cache_data
def loading_data():
    df = load_data("Sheet1")
    df2 = load_data("m_branch")
    df3 = load_data("m_source")
    df4 = load_data("m_vendor")
    return df, df2, df3, df4

df, df2, df3, df4 = loading_data()

def append_to_database(new_data):
    sheet_db = connect_gsheet("Sheet1")
    for _, row in new_data.iterrows():
        sheet_db.append_row([
            row["Branch"],
            row["Source"],
            str(row["Tanggal Terima"].strftime("%d/%m/%Y")),
            row["Vendor"],
            row["No Invoice"],
            row["Keterangan"],
            row["Nominal"],
            str(row["Payment Date"].strftime("%d/%m/%Y")),
            row["Payment Code"],
            row.get("Invoice", "Not Done"),
            row.get("FP", "Not Done"),
            row.get("Surat Jalan", "Not Done"),
            row.get("Refund", "No Refund"),
            row.get("Notes", ""),
            row.get("Nomor Ticket", "")
        ])

st.title("Input Data")

if st.button("Reload Data"):
    time.sleep(1)
    st.cache_data.clear()
    df, df2, df3, df4 = loading_data()
    st.rerun()

Branch = st.selectbox("Branch", df2['Branch'].sort_values().unique())
Source = st.selectbox("Source", df3['Source'].sort_values().unique())
Date = st.date_input("Tanggal Terima")
Vendor = st.selectbox("Vendor", df4['BP Name'].sort_values().unique())
No_Invoice = st.text_input("No Invoice")
Keterangan = st.text_input("Keterangan")
Price = st.number_input("Nominal Rp.", min_value=0)
Date2 = st.date_input("Tanggal Payment")
Ticket = st.text_input("Nomor Ticket")

prefix = f"{Branch}/{yyyymm}/"
df_live = load_data("Sheet1")
existing = df_live[df_live["Payment Code"].astype(str).str.startswith(prefix, na=False)]
if existing.empty:
    seq = 1
else:
    last_num = (
        existing["Payment Code"]
        .astype(str)
        .str.extract(r"(\d{4})$")[0]
        .astype(int)
        .max()
    )
    seq = last_num + 1
Code = f"{prefix}{seq:04d}"

df_live = load_data("Sheet1")
df_live["Payment Date"] = pd.to_datetime(df_live["Payment Date"], errors="coerce")

def make_key(branch, source, vendor, invoice, date):
    return f"{branch}|{source}|{vendor}|{invoice}|{pd.to_datetime(date)}"

existing_keys = df_live.apply(
    lambda r: make_key(
        r["Branch"],
        r["Source"],
        r["Vendor"],
        r["No Invoice"],
        r["Payment Date"]
    ),
    axis=1
).astype(str)

current_key = make_key(Branch, Source, Vendor, No_Invoice, Date2)

is_duplicate = current_key in set(existing_keys)

# 🔥 LIVE WARNING (shows before submit)
if is_duplicate:
    st.warning("⚠️ Duplicate detected! This combination already exists.")

# 🔥 DISABLE BUTTON BEFORE SUBMIT
if st.button("Submit", disabled=is_duplicate):
    new_data = pd.DataFrame({
        "Branch": [Branch],
        "Source": [Source],
        "Tanggal Terima": [Date],
        "Vendor": [Vendor],
        "No Invoice": [No_Invoice],
        "Keterangan": [Keterangan],
        "Nominal": [Price],
        "Payment Date": [Date2],
        "Payment Code": [Code],
        "Invoice": ["Not Done"],
        "FP": ["Not Done"],
        "Surat Jalan": ["Not Done"],
        "Refund": ["No Refund"],
        "Notes": [""],
        "Nomor Ticket": [Ticket]
    })

    append_to_database(new_data)
    st.success("Data submitted successfully!")
    st.cache_data.clear()
    st.cache_resource.clear()
    st.switch_page("Summary.py")
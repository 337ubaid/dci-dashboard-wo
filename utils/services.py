import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

@st.cache_resource
def get_client():
    """
    Create and return an authenticated Google Sheets client using gspread.
    Return:
    -------
    gspread.Client
        An authorized gspread client instance connected with the given 
        service account credentials.
    """
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds)


def get_worksheet(link_spreadsheet=None, nama_worksheet="DATABASE"):
    """
    Ambil worksheet tertentu dari Google Spreadsheet.

    Param:
        - link_spreadsheet (str): URL Spreadsheet (default ambil dari st.session_state)
        - nama_worksheet (str): nama tab worksheet (default "DATABASE")
    Return:
        - worksheet object
    """
    if link_spreadsheet is None:
        link_spreadsheet = st.session_state.get("database_gsheet_url", "")
    if nama_worksheet is None:
        nama_worksheet = st.session_state.get("database_sheet_name", "DATABASE")

    if not link_spreadsheet:
        raise ValueError("❌ Link spreadsheet tidak ditemukan.")
    
    client = get_client()
    return client.open_by_url(link_spreadsheet).worksheet(nama_worksheet)

def get_raw_values(link_spreadsheet=None, nama_worksheet="DATABASE", usecols=None):
    """
    Ambil nilai mentah dari worksheet Google Sheets.

    Param:
        - link_spreadsheet (str): URL Spreadsheet (default ambil dari st.session_state)
        - nama_worksheet (str): nama tab worksheet (default "DATABASE")
        - usecols (list): daftar kolom yang mau diambil (default None = semua)
    Return:
        - DataFrame dengan data mentah dari worksheet
    """

    worksheet_name = get_worksheet(link_spreadsheet, nama_worksheet)
    raw_all_values = worksheet_name.get_all_values()

    if not raw_all_values:
        st.warning(f"Sheet {nama_worksheet} kosong.")
        return pd.DataFrame()

    # header baris 6 (mulai B)
    header = raw_all_values[5][1:]
    # values mulai baris 8 (mulai B)
    values = [row[1:] for row in raw_all_values[7:]]

    if not values:
        st.warning(f"Sheet {nama_worksheet} tidak memiliki data.")
        return pd.DataFrame(columns=header)

    df = pd.DataFrame(values, columns=header)

    # filter kolom kalau ada usecols
    if usecols:
        missing = [col for col in usecols if col not in df.columns]
        if missing:
            st.warning(f"Kolom tidak ditemukan: {missing}")
        df = df[[col for col in usecols if col in df.columns]]

    df.reset_index(drop=True, inplace=True)
    return df

import streamlit as st
import pandas as pd
import re
import gspread
from gspread.utils import ValueRenderOption
from google.oauth2.service_account import Credentials

# =========================
# AUTENTIKASI
# =========================
@st.cache_resource
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds)


def get_worksheet(link_spreadsheet, nama_worksheet):
    client = get_client()
    return client.open_by_url(link_spreadsheet).worksheet(nama_worksheet)

# =========================
# AMBIL DATA TREMS per kolom
# =========================
worksheet = get_worksheet(st.secrets["spreadsheet_database"]["spreadsheet_link"], "trems")


def get_column_as_df(worksheet, col_letter, column_name, render_option=ValueRenderOption.unformatted):
    """
    Ambil 1 kolom dari Google Sheet, return DataFrame.
    
    Args:
        worksheet: gspread worksheet object
        col_letter: str → huruf kolom, misalnya "C"
        start_row: int → baris mulai (misal 8)
        end_row: int → baris akhir (misal worksheet.row_count-4)
        column_name: str → nama kolom DataFrame
        render_option: gspread.utils.ValueRenderOption (default unformatted)
    """
    values = worksheet.get(
        f"{col_letter}8:{col_letter}",
        value_render_option=render_option
    )
    return pd.DataFrame(values, columns=[column_name])


# ambil kolom yang diperlukan
df_no_jastel = get_column_as_df(worksheet, "C", "No.Jastel")
df_bill_periode = get_column_as_df(worksheet, "H", "Bill.Pe")
df_payment_date = get_column_as_df(worksheet, "I", "Payment Dat", render_option=ValueRenderOption.formatted)
df_collection_agent = get_column_as_df(worksheet, "J", "Collection Agent")
df_total_amount = get_column_as_df(worksheet, "M", "Total Amount")

st.write(df_no_jastel.head(), df_no_jastel.shape)
st.write(df_bill_periode.head(), df_bill_periode.shape)
st.write(df_payment_date.head(), df_payment_date.shape)
st.write(df_collection_agent.head(), df_collection_agent.shape)
st.write(df_total_amount.head(), df_total_amount.shape)

# gabung semua kolom
df_trems = pd.concat(
    [df_no_jastel, 
     df_bill_periode, 
     df_payment_date, 
     df_collection_agent, 
     df_total_amount],
    axis=1
).reset_index(drop=True)


# format payment date ke dd/mm/yyyy
df_trems["Payment Dat"] = pd.to_datetime(df_trems["Payment Dat"]).dt.strftime("%d/%m/%Y")

st.dataframe(df_trems)

# hapus payment date kosong
df_trems = df_trems[df_trems["Bill.Pe"].notna() & (df_trems["Bill.Pe"] != "")]

# cek tipe jastel
# fungsi mapping nomor jastel ke kategori
def kategori_jastel(no):
    no_str = str(no)
    if no_str.startswith("1"):
        return "INET"
    elif no_str.startswith("3"):
        return "TELP"
    elif no_str.startswith("7"):
        return "DIGIPRO"
    else:
        return "LAINNYA"

# tambahkan kolom kategori ke dataframe utama
df_trems["Kategori"] = df_trems["No.Jastel"].astype(str).apply(kategori_jastel)

st.dataframe(df_trems)

# kelompokkan berdasarkan bill periode menjadi dataframe terpisah
# ambil dua nilai Bill Periode yang ada
periode1, periode2 = pd.to_numeric(df_trems["Bill.Pe"].unique()).astype(int)
st.write(f"Period 1: {periode1}, Period 2: {periode2}")


# =========================
# buat dataframe per periode
# =========================
df_periode1 = df_trems[df_trems["Bill.Pe"] == periode1].reset_index(drop=True)
df_periode2 = df_trems[df_trems["Bill.Pe"] == periode2].reset_index(drop=True)

# ubah tipe data menjadi integer / numeric
for df in [df_periode1, df_periode2]:
    df["No.Jastel"] = pd.to_numeric(df["No.Jastel"], errors="coerce").fillna(0).astype(int)
    df["Bill.Pe"] = pd.to_numeric(df["Bill.Pe"], errors="coerce").fillna(0).astype(int)

    for col in ["Payment Dat", "Collection Agent"]:
        df[col] = df[col].fillna("").astype(str)
st.write("Data Periode 1:")
st.dataframe(df_periode1)
st.write("Data Periode 2:")
st.dataframe(df_periode2)


# =========================
# fungsi worksheet helper
# =========================
def get_or_create_worksheet(spreadsheet, title, rows, cols):
    import gspread
    try:
        ws = spreadsheet.worksheet(title)
        ws.clear()  # reset isinya
        return ws
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


# =========================
# buat / ambil worksheet sesuai periode
# =========================
worksheet_periode1 = get_or_create_worksheet(
    worksheet.spreadsheet,
    f"trems_{period1}",
    rows=max(len(df_periode1)+1, 100),
    cols=df_periode1.shape[1]
)

worksheet_periode2 = get_or_create_worksheet(
    worksheet.spreadsheet,
    f"trems_{period2}",
    rows=max(len(df_periode2)+1, 100),
    cols=df_periode2.shape[1]
)


# =========================
# siapkan data untuk update (header + isi)
# =========================
values1 = [df_periode1.columns.tolist()] + df_periode1.astype(str).values.tolist()
values2 = [df_periode2.columns.tolist()] + df_periode2.astype(str).values.tolist()


# =========================
# update ke worksheet
# =========================
worksheet_periode1.update("A1", values1)
worksheet_periode2.update("A1", values2)

# cek nomor jastel, jika diawali 3 maka tipe INET, jika 8 maka tipe POTS, jika 7 maka digital
def cek_tipe_jastel(no_jastel):
    str_jastel = str(no_jastel)
    if str_jastel.startswith("3"):
        return "INET"
    elif str_jastel.startswith("8"):
        return "POTS"
    elif str_jastel.startswith("7"):
        return "DIGITAL"
    else:
        return "UNKNOWN"
 
# =========================
# Ambil data dari sheet MASTER
# =========================
MASTER_SHEET_NAME = "tes database"
worksheet_master = get_worksheet(st.secrets["spreadsheet_database"]["spreadsheet_link"], MASTER_SHEET_NAME)

data_master = worksheet_master.get_all_values()
# st.dataframe(data_master)
if not data_master:
    st.error(f"Sheet '{MASTER_SHEET_NAME}' kosong atau tidak ditemukan.")
else:
    # jadikan DataFrame
    df_master = pd.DataFrame(data_master[1:], columns=data_master[0])
    # st.dataframe(df_master)

    # ambil hanya kolom INET, TELP, DIGIPRO (kalau ada)
    target_cols = ["INET", "TELP", "DIGIPRO"]
    df_jastel_master = df_master[[c for c in target_cols if c in df_master.columns]]

    # tampilkan
    st.subheader("Kolom Jastel dari MASTER DATABASE")
    st.dataframe(df_jastel_master)

# normalize helper: buang karakter non-digit sehingga format konsisten
def normalize_jastel(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    if s in ["", "-"]:
        return ""
    return re.sub(r"\D+", "", s)

# 1) build lookup dari MASTER: nomor_jastel -> (row_index, kategori_column_name)
jastel_cols = ["INET", "TELP", "DIGIPRO"]
master_lookup = {}
for idx, row in df_jastel_master.iterrows():
    for cat in jastel_cols:
        cell = row.get(cat, "")
        if pd.isna(cell) or str(cell).strip() in ["", "-"]:
            continue
        # split jika ada beberapa nomor (pisah koma / ; / spasi)
        parts = re.split(r"[,\n;]+", str(cell))
        for p in parts:
            p_norm = normalize_jastel(p)
            if p_norm:
                master_lookup[p_norm] = (idx, cat)

# 2) buat kolom SAP/TGL/TAG untuk suffix period (bisa panggil ulang untuk periode lain)
def ensure_period_cols(df_master, suffix, jastel_cols=jastel_cols):
    prefixes = ["SAP", "TGL", "TAG"]
    for cat in jastel_cols:
        for pre in prefixes:
            col = f"{pre} {cat}{suffix}"
            if col not in df_master.columns:
                df_master[col] = ""   # buat kolom baru tetap mempertahankan urutan kolom MASTER lain
    return df_master

def update_master_with_period(df_master, df_period, master_lookup, suffix):
    jastel_cols = ["INET", "TELP", "DIGIPRO"]
    df_master = ensure_period_cols(df_master, suffix, jastel_cols=jastel_cols)

    update_logs = []
    for _, r in df_period.iterrows():
        raw_nj = r.get("No.Jastel", "")
        nj = normalize_jastel(raw_nj)
        if not nj:
            continue

        hit = master_lookup.get(nj)
        if not hit:
            update_logs.append({
                "No.Jastel": raw_nj, "Normalized": nj, "Action": "NOT_FOUND_IN_MASTER"
            })
            continue

        idx, matched_cat = hit
        sap_col = f"SAP {matched_cat}{suffix}"
        tgl_col = f"TGL {matched_cat}{suffix}"
        tag_col = f"TAG {matched_cat}{suffix}"

        new_sap = str(r.get("Collection Agent", "")).strip()
        new_tgl_raw = r.get("Payment Dat", "")
        new_tag = r.get("Total Amount", "")

        # format tanggal
        new_tgl = ""
        try:
            if pd.notna(new_tgl_raw) and str(new_tgl_raw).strip() not in ["", "-"]:
                new_tgl = pd.to_datetime(new_tgl_raw, dayfirst=True).strftime("%d/%m/%Y")
        except Exception:
            new_tgl = str(new_tgl_raw).strip()

        # ambil existing value
        old_sap = str(df_master.at[idx, sap_col]) if sap_col in df_master.columns else ""
        old_tgl = str(df_master.at[idx, tgl_col]) if tgl_col in df_master.columns else ""
        old_tag = str(df_master.at[idx, tag_col]) if tag_col in df_master.columns else ""

        # update jika beda
        if new_sap and new_sap != old_sap:
            df_master.at[idx, sap_col] = new_sap
            update_logs.append({"No.Jastel": raw_nj, "Kategori": matched_cat,
                                "Kolom": sap_col, "Old": old_sap, "New": new_sap})

        if new_tgl and new_tgl != old_tgl:
            df_master.at[idx, tgl_col] = new_tgl
            update_logs.append({"No.Jastel": raw_nj, "Kategori": matched_cat,
                                "Kolom": tgl_col, "Old": old_tgl, "New": new_tgl})

        if (pd.notna(new_tag) and str(new_tag).strip()) and str(new_tag) != old_tag:
            df_master.at[idx, tag_col] = new_tag
            update_logs.append({"No.Jastel": raw_nj, "Kategori": matched_cat,
                                "Kolom": tag_col, "Old": old_tag, "New": new_tag})

    return df_master, update_logs

# Update periode 1
suffix1 = str(period1)[-2:]
df_master, logs1 = update_master_with_period(df_master, df_periode1, master_lookup, suffix1)

# Update periode 2
suffix2 = str(period2)[-2:]
df_master, logs2 = update_master_with_period(df_master, df_periode2, master_lookup, suffix2)

# Gabung log
all_logs = logs1 + logs2

# Preview hasil akhir
sap_tgl_tag_cols = [c for c in df_master.columns if re.match(r"^(SAP|TGL|TAG)\s", c)]
cols_preview = jastel_cols + sap_tgl_tag_cols
st.subheader("Preview MASTER setelah update periode 1 & 2")
st.dataframe(df_master[cols_preview])

st.subheader("Log update (sample)")
st.dataframe(pd.DataFrame(all_logs))


# # 5) tombol commit (jika kamu yakin ingin push ke sheet MASTER)
# if st.button(f"Commit perubahan ke sheet MASTER (periode {period1})"):
#     try:
#         values = [df_master.columns.tolist()] + df_master.fillna("").values.tolist()
#         worksheet_master.update("A1", values)
#         st.success("Master sheet updated successfully (SAP/TGL/TAG columns only).")
#     except Exception as e:
#         st.error(f"Gagal update master sheet: {e}")
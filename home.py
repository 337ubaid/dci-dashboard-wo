import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

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
# AMBIL DATA TREMS
# =========================
def load_trems_df(sheet):
    """
    Ambil header di B6 ke kanan, dan values mulai B8 ke kanan + bawah.
    """
    all_values = sheet.get_all_values()

    # Header di baris ke-6, mulai kolom B (index 1 → Python 0-based)
    header = all_values[5][1:]  

    # Data mulai baris ke-8, kolom B ke kanan
    values = [row[1:] for row in all_values[7:]]

    df = pd.DataFrame(values, columns=header)
    return df


# =========================
# PROSES DATA TREMS
# =========================
def proses_trems(df):
    # kolom indeks sesuai script AppScript
    idxNoJastel = "No.Jastel"
    idxBillPeriode = "Bill.Pe"
    idxPaymentDate = "Payment Dat"
    idxCollAgent = "Collection Agent"
    idxTotalAmount = "Total Amount"

    st.dataframe(df.head())

    # Step 1: Filter
    df = df[(df[idxBillPeriode] != "0") & (df[idxBillPeriode] != "") & (df[idxNoJastel] != "")]

    # Step 2: Collection Agent kosong → UNPAID
    df[idxCollAgent] = df[idxCollAgent].replace("", "UNPAID")

    # Step 3: Hapus duplikat (ambil payment date terbaru kalau ada)
    df[idxPaymentDate] = pd.to_datetime(df[idxPaymentDate], errors="coerce")
    df = df.sort_values(idxPaymentDate, ascending=False).drop_duplicates(
        [idxNoJastel, idxBillPeriode], keep="first"
    )

    # Step 4: Payment Date ada tapi CollAgent = UNPAID → ganti Billing Nol
    mask = df[idxPaymentDate].notna() & (df[idxCollAgent].str.upper() == "UNPAID")
    df.loc[mask, idxCollAgent] = "Billing Nol"

    return df


# =========================
# UPDATE WO REG
# =========================
def update_wo_reg(df_by_period, sheetWO):
    data_wo = sheetWO.get_all_values()
    headers, rows = data_wo[0], data_wo[1:]
    df_wo = pd.DataFrame(rows, columns=headers)

    for period, rows in df_by_period.items():
        suffix = str(period)[-2:]  # ambil 2 digit bulan

        idxSapInet = f"SAP INET{suffix}"
        idxSapTlp = f"SAP TLP{suffix}"
        idxSapDig = f"SAP DIG{suffix}"
        idxTglInet = f"TGL INET{suffix}"
        idxTglTlp = f"TGL TLP{suffix}"
        idxTglDig = f"TGL DIG{suffix}"
        idxTagInet = f"TAG INET{suffix}"
        idxTagTlp = f"TAG TLP{suffix}"
        idxTagDig = f"TAG DIG{suffix}"

        mapTrems = {r[0]: {"payDate": r[2], "coll": r[3], "total": r[4]} for r in rows}

        # Loop row WO
        for i, row in df_wo.iterrows():
            inet, telp, digi = row[6], row[7], row[8]

            # INET
            if inet and inet != "-" and inet in mapTrems:
                df_wo.at[i, idxSapInet] = mapTrems[inet]["coll"]
                df_wo.at[i, idxTglInet] = mapTrems[inet]["payDate"]
                df_wo.at[i, idxTagInet] = mapTrems[inet]["total"]
            elif inet == "-":
                df_wo.at[i, idxSapInet] = "NO BILL"
                df_wo.at[i, idxTglInet] = "-"
                df_wo.at[i, idxTagInet] = "0"

            # TELP
            if telp and telp != "-" and telp in mapTrems:
                df_wo.at[i, idxSapTlp] = mapTrems[telp]["coll"]
                df_wo.at[i, idxTglTlp] = mapTrems[telp]["payDate"]
                df_wo.at[i, idxTagTlp] = mapTrems[telp]["total"]
            elif telp == "-":
                df_wo.at[i, idxSapTlp] = "NO BILL"
                df_wo.at[i, idxTglTlp] = "-"
                df_wo.at[i, idxTagTlp] = "0"

            # DIGI
            if digi and digi != "-" and digi in mapTrems:
                df_wo.at[i, idxSapDig] = mapTrems[digi]["coll"]
                df_wo.at[i, idxTglDig] = mapTrems[digi]["payDate"]
                df_wo.at[i, idxTagDig] = mapTrems[digi]["total"]
            elif digi == "-":
                df_wo.at[i, idxSapDig] = "NO BILL"
                df_wo.at[i, idxTglDig] = "-"
                df_wo.at[i, idxTagDig] = "0"

    # 🚀 Update balik ke sheet tanpa hapus seluruh sheet
    # Hanya tulis ulang DataFrame hasil update ke kolom yang diubah
    sheetWO.update([headers] + df_wo.values.tolist())


# =========================
# MAIN STREAMLIT
# =========================
def main():
    st.title("🔄 Proses Data Trems via Streamlit")

    url = st.text_input("Spreadsheet URL:", st.session_state.get("database_gsheet_url", ""))

    if st.button("Jalankan Proses") and url:
        client = get_client()
        sheetTrems = get_worksheet(url, "trems")
        sheetWO = get_worksheet(url, "WO REG")

        # Load & proses data
        df_trems = load_trems_df(sheetTrems)
        df_clean = proses_trems(df_trems)

        # Bagi berdasarkan periode
        df_by_period = {
            period: group[[ "No.Jastel", "Bill.Pe", "Payment Dat", "Collection Agent", "Total Amount"]].values.tolist()
            for period, group in df_clean.groupby("Bill.Pe")
        }

        # Update WO REG
        update_wo_reg(df_by_period, sheetWO)

        st.success("Proses selesai ✅")
        st.write("Periode diproses:", list(df_by_period.keys()))


if __name__ == "__main__":
    main()

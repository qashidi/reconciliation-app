
import pandas as pd
import streamlit as st
from datetime import timedelta
import os

# --- CONFIGURABLE PARAMETERS ---
DATE_TOLERANCE_DAYS = 2
AMOUNT_TOLERANCE_PERCENT = 0.01  # 1%

# --- FUNCTIONS ---
def load_data(sales_file, bank_file) -> tuple[pd.DataFrame, pd.DataFrame]:
    sales_data = pd.read_excel(sales_file)
    bank_data = pd.read_excel(bank_file)
    return sales_data, bank_data

def preprocess_data(sales_data: pd.DataFrame, bank_data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    sales_data['Tanggal'] = pd.to_datetime(sales_data['Tanggal'])
    bank_data['Tanggal'] = pd.to_datetime(bank_data['Tanggal'])
    sales_data['Jumlah'] = sales_data['Jumlah'].astype(float)
    bank_data['Jumlah'] = bank_data['Jumlah'].astype(float)
    return sales_data, bank_data

def reconcile(sales_data: pd.DataFrame, bank_data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    matched = []
    unmatched_sales = sales_data.copy()
    unmatched_bank = bank_data.copy()

    for idx, sale in sales_data.iterrows():
        possible_matches = bank_data[
            (bank_data['Jumlah'].between(
                sale['Jumlah'] * (1 - AMOUNT_TOLERANCE_PERCENT),
                sale['Jumlah'] * (1 + AMOUNT_TOLERANCE_PERCENT))) &
            (abs(bank_data['Tanggal'] - sale['Tanggal']) <= timedelta(days=DATE_TOLERANCE_DAYS))
        ]

        if not possible_matches.empty:
            matched_bank = possible_matches.iloc[0]
            matched.append({
                'Invoice': sale['Invoice'],
                'Sales_Date': sale['Tanggal'],
                'Sales_Amount': sale['Jumlah'],
                'Bank_Date': matched_bank['Tanggal'],
                'Bank_Amount': matched_bank['Jumlah']
            })

            unmatched_sales = unmatched_sales[unmatched_sales['Invoice'] != sale['Invoice']]
            unmatched_bank = unmatched_bank[unmatched_bank.index != matched_bank.name]

    matched_df = pd.DataFrame(matched)
    return matched_df, unmatched_sales, unmatched_bank

def export_results(matched: pd.DataFrame, unmatched_sales: pd.DataFrame, unmatched_bank: pd.DataFrame, output_file: str):
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    with pd.ExcelWriter(output_file) as writer:
        matched.to_excel(writer, sheet_name='Matched', index=False)
        unmatched_sales.to_excel(writer, sheet_name='Unmatched_Sales', index=False)
        unmatched_bank.to_excel(writer, sheet_name='Unmatched_Bank', index=False)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Rekonsiliasi Produksi vs Kas", layout="wide")
st.title("🧾 Aplikasi Rekonsiliasi Produksi vs Penerimaan Kas")

st.sidebar.header("Upload Data")
sales_file = st.sidebar.file_uploader("Upload file Data Penjualan (Excel)", type=["xlsx"])
bank_file = st.sidebar.file_uploader("Upload file Data Rekening Koran (Excel)", type=["xlsx"])

if sales_file and bank_file:
    sales_data, bank_data = load_data(sales_file, bank_file)
    sales_data, bank_data = preprocess_data(sales_data, bank_data)

    matched, unmatched_sales, unmatched_bank = reconcile(sales_data, bank_data)

    st.success("🎉 Rekonsiliasi Selesai!")
    st.subheader("Data Matched")
    st.dataframe(matched, use_container_width=True)

    st.subheader("Sales Tidak Ada Penerimaan")
    st.dataframe(unmatched_sales, use_container_width=True)

    st.subheader("Penerimaan Tidak Ada Penjualan")
    st.dataframe(unmatched_bank, use_container_width=True)

    with st.spinner('Mempersiapkan file untuk diunduh...'):
        output_path = 'output/reconciliation_result.xlsx'
        export_results(matched, unmatched_sales, unmatched_bank, output_path)
        with open(output_path, "rb") as file:
            st.download_button(
                label="📥 Download File Rekonsiliasi",
                data=file,
                file_name="reconciliation_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("📂 Silakan upload kedua file untuk memulai rekonsiliasi.")

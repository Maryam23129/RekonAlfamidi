import streamlit as st
import pandas as pd
import re
from io import BytesIO

def load_excel(file):
    return pd.read_excel(file)

def extract_total_rekening(df):
    df = df.iloc[12:, [1, 2, 5]].dropna()
    df.columns = ['Tanggal', 'Remark', 'Credit']
    df = df[df['Remark'].str.contains("DARI MIDI UTAMA INDONESIA", case=False, na=False)]
    df['Credit'] = df['Credit'].replace('[^0-9.]', '', regex=True).astype(float)
    df['KodeTanggal'] = df['Remark'].str.extract(r'^(\S+)')[0].str[-4:]
    df['Bulan'] = df['KodeTanggal'].str[:2]
    df['Tanggal'] = df['KodeTanggal'].str[2:]
    df['Tanggal Transaksi'] = pd.to_datetime('2025' + df['Bulan'] + df['Tanggal'], format='%Y%m%d', errors='coerce')
    return df

st.set_page_config(page_title="Dashboard Rekonsiliasi Pendapatan Ticketing", layout="wide")

st.markdown("""
<h1 style='text-align: center;'>📊 Dashboard Rekonsiliasi Pendapatan Ticketing 🚢💰</h1>
<p style='text-align: center; font-size: 18px;'>Bandingkan data invoice dan uang masuk dari rekening dengan mudah.</p>
""", unsafe_allow_html=True)

# Upload
uploaded_invoice = st.sidebar.file_uploader("📥 Upload File Invoice", type=["xlsx"])
uploaded_rekening = st.sidebar.file_uploader("📥 Upload File Rekening Koran", type=["xlsx"])

if uploaded_invoice and uploaded_rekening:
    # Load data
    invoice_df = load_excel(uploaded_invoice)
    rekening_df = load_excel(uploaded_rekening)
    rekening_detail_df = extract_total_rekening(rekening_df)

    st.subheader("📄 Tabel Rekonsiliasi Invoice dan Rekening Koran per Tanggal")

    invoice_df['TANGGAL INVOICE'] = pd.to_datetime(invoice_df['TANGGAL INVOICE'], errors='coerce')
    invoice_df['HARGA'] = pd.to_numeric(invoice_df['HARGA'], errors='coerce')

    # Rekap invoice harian dari file invoice
    rekap_invoice = (
        invoice_df
        .groupby(invoice_df['TANGGAL INVOICE'].dt.strftime('%d-%m-%Y'))['HARGA']
        .sum()
        .reset_index()
        .rename(columns={'TANGGAL INVOICE': 'Tanggal Transaksi', 'HARGA': 'Total Invoice'})
    )

    # Rekap rekening (0066)
    rekening_0066 = rekening_detail_df[rekening_detail_df['Remark'].str.contains("0066", case=False, na=False)].copy()
    rekening_0066['Tanggal Transaksi'] = rekening_0066['Tanggal Transaksi'].dt.strftime('%d-%m-%Y')

    rekap_rekening = (
        rekening_0066
        .groupby('Tanggal Transaksi')['Credit']
        .sum()
        .reset_index()
        .rename(columns={'Credit': 'Uang Masuk'})
    )

    rekap_final = pd.merge(rekap_invoice, rekap_rekening, on='Tanggal Transaksi', how='outer').fillna(0)
    rekap_final['Selisih'] = rekap_final['Total Invoice'] - rekap_final['Uang Masuk']

    st.dataframe(rekap_final, use_container_width=True)

else:
    st.info("Silakan upload file Invoice dan Rekening Koran untuk memulai.")

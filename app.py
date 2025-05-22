import streamlit as st
import pandas as pd
from io import BytesIO
import re

def load_excel(file):
    return pd.read_excel(file)

def extract_total_rekening(df):
    df = df.iloc[12:, [1, 2, 5]].dropna()
    df.columns = ['Tanggal', 'Remark', 'Credit']
    df = df[df['Remark'].str.contains("DARI MIDI UTAMA INDONESIA", case=False, na=False)]
    df['Credit'] = df['Credit'].replace('[^0-9.]', '', regex=True).astype(float)
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
    output.seek(0)
    return output

st.set_page_config("Rekonsiliasi Tiket", layout="wide")
st.title("📊 Dashboard Rekonsiliasi Pendapatan Ticketing")

uploaded = st.sidebar.file_uploader("📁 Upload Semua File Excel", type="xlsx", accept_multiple_files=True)

tiket_files, invoice_file, rekening_file, summary_files = [], None, None, []

for f in uploaded:
    name = f.name.lower()
    if "tiket" in name:
        tiket_files.append(f)
    elif "invoice" in name:
        invoice_file = f
    elif "rekening" in name:
        rekening_file = f
    elif "summary" in name:
        summary_files.append(f)

if tiket_files and invoice_file and rekening_file and summary_files:
    b2b_data = []
    for tf in tiket_files:
        df = load_excel(tf)
        row = df[df.apply(lambda r: r.astype(str).str.contains("TOTAL JUMLAH", case=False).any(), axis=1)]
        if not row.empty:
            pelabuhan = next((p for p in ["merak", "bakauheni", "ketapang", "gilimanuk", "ciwandan", "panjang"] if p in tf.name.lower()), "lainnya")
            b2b_data.append({"Pelabuhan": pelabuhan.capitalize(), "Pendapatan": pd.to_numeric(row.iloc[0, 4], errors='coerce')})

    invoice_df = load_excel(invoice_file)
    invoice_df["HARGA"] = pd.to_numeric(invoice_df["HARGA"], errors="coerce")
    invoice_df = invoice_df[invoice_df["STATUS"].str.lower() == "dibayar"]

    rekening_df = load_excel(rekening_file)
    rekening_detail = extract_total_rekening(rekening_df)
    total_rekening = rekening_detail["Credit"].sum()

    penambahan_dict = {}
    pengurangan_total = ""
    naik_turun_dict = {}

    match = re.search(r'(\d{4}-\d{2}-\d{2})\s*s[_\-]d\s*(\d{4}-\d{2}-\d{2})', invoice_file.name)
    if match:
        start_str, end_str = match.groups()
        start_date, end_date = pd.to_datetime(start_str), pd.to_datetime(end_str)
        tanggal_display = f"{start_date:%d-%m-%Y} s.d {end_date:%d-%m-%Y}"

        invoice_df["NOMOR INVOICE"] = invoice_df.get("NOMER INVOICE", invoice_df.get("NOMOR INVOICE", "")).astype(str)

        for sf in summary_files:
            if start_str in sf.name:
                df = load_excel(sf)
                if "NOMOR INVOICE" in df.columns and "ASAL" in df.columns:
                    df["NOMOR INVOICE"] = df["NOMOR INVOICE"].astype(str)
                    df["ASAL"] = df["ASAL"].astype(str).str.lower()
                    for pel in ["merak", "bakauheni", "ketapang", "gilimanuk", "ciwandan", "panjang"]:
                        sum_pel = df[df["ASAL"] == pel]
                        inv_pel = invoice_df[invoice_df["KEBERANGKATAN"].str.lower().str.contains(pel)]
                        ringkasan = []
                        common_invoices = set(sum_pel["NOMOR INVOICE"]).intersection(set(inv_pel["NOMOR INVOICE"]))
                        for noinv in common_invoices:
                            t1 = sum_pel[sum_pel["NOMOR INVOICE"] == noinv]["TARIF"].sum()
                            t2 = inv_pel[inv_pel["NOMOR INVOICE"] == noinv]["HARGA"].sum()
                            if t1 != t2:
                                ringkasan.append(f"{noinv}: S={int(t1)}, I={int(t2)}")
                        if ringkasan:
                            naik_turun_dict[pel] = "; ".join(ringkasan)
                break

        st.write("🧾 Log Naik Turun Golongan:", naik_turun_dict)

    pelabuhans = ["Merak", "Bakauheni", "Ketapang", "Gilimanuk", "Ciwandan", "Panjang"]
    invoice_vals = [invoice_df[invoice_df["KEBERANGKATAN"].str.lower().str.contains(p.lower())]["HARGA"].sum() for p in pelabuhans]
    uang_masuk_vals = [total_rekening] + [0]*5
    selisih_vals = [i - u for i, u in zip(invoice_vals, uang_masuk_vals)]

    df = pd.DataFrame({
        "No": range(1, 7),
        "Tanggal Transaksi": [tanggal_display]*6,
        "Pelabuhan Asal": pelabuhans,
        "Nominal Tiket Terjual": [next((b["Pendapatan"] for b in b2b_data if b["Pelabuhan"].lower() == p.lower()), 0) for p in pelabuhans],
        "Invoice": invoice_vals,
        "Uang Masuk": uang_masuk_vals,
        "Selisih": selisih_vals,
        "Pengurangan": [pengurangan_total if i == 0 else "" for i in range(6)],
        "Penambahan": [penambahan_dict.get(p.lower(), "") for p in pelabuhans],
        "Naik Turun Golongan": [naik_turun_dict.get(p.lower(), "") for p in pelabuhans],
        "NET": ["" for _ in pelabuhans]
    })

    total = {
        "No": "", "Tanggal Transaksi": "", "Pelabuhan Asal": "TOTAL",
        "Nominal Tiket Terjual": df["Nominal Tiket Terjual"].sum(),
        "Invoice": df["Invoice"].sum(),
        "Uang Masuk": df["Uang Masuk"].sum(),
        "Selisih": df["Selisih"].sum(),
        "Pengurangan": "", "Penambahan": "", "Naik Turun Golongan": "", "NET": ""
    }
    df = pd.concat([df, pd.DataFrame([total])], ignore_index=True)

    st.subheader("📄 Rekap Per Pelabuhan")
    st.dataframe(
        df[df["Pelabuhan Asal"] != "TOTAL"].drop(columns=["Invoice", "Uang Masuk", "Selisih"]),
        use_container_width=True
    )

    st.subheader("📄 Tabel Rekonsiliasi Invoice dan Rekening Koran")
    st.dataframe(
        df[df["Pelabuhan Asal"] == "TOTAL"].drop(columns=[
            "Pelabuhan Asal", "No", "Nominal Tiket Terjual",
            "Pengurangan", "Penambahan", "Naik Turun Golongan", "NET"
        ]),
        use_container_width=True
    )

    st.download_button("📥 Download Excel", to_excel(df), file_name="rekapitulasi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("📂 Silakan upload semua file: tiket, invoice, summary, rekening.")

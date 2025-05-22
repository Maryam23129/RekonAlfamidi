import streamlit as st
import pandas as pd
from io import BytesIO
import re

def load_excel(file):
    return pd.read_excel(file)

def extract_total_rekening(rekening_df):
    rekening_df = rekening_df.iloc[12:, [1, 2, 5]].dropna()
    rekening_df.columns = ['Tanggal', 'Remark', 'Credit']
    rekening_df = rekening_df[rekening_df['Remark'].str.contains("DARI MIDI UTAMA INDONESIA", case=False, na=False)]
    rekening_df['Credit'] = rekening_df['Credit'].replace('[^0-9.]', '', regex=True).astype(float)
    rekening_df['TanggalKode'] = rekening_df['Remark'].str.extract(r'^(\S+)')[0].str[-4:]
    rekening_df['Bulan'] = rekening_df['TanggalKode'].str[:2]
    rekening_df['Tanggal'] = rekening_df['TanggalKode'].str[2:]
    rekening_df['Tanggal Transaksi'] = pd.to_datetime('2025' + rekening_df['Bulan'] + rekening_df['Tanggal'], format='%Y%m%d', errors='coerce')
    return rekening_df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
        workbook = writer.book
        worksheet = writer.sheets['Rekapitulasi']
        currency_format = workbook.add_format({'num_format': '"Rp" #,##0'})
        for col_num, column in enumerate(df.columns):
            fmt = currency_format if column in ['Nominal Tiket Terjual', 'Invoice', 'Uang Masuk', 'Selisih'] else None
            worksheet.set_column(col_num, col_num, 20, fmt)
    output.seek(0)
    return output

st.set_page_config(page_title="Rekonsiliasi Tiket", layout="wide")

st.title("📊 Rekonsiliasi Pendapatan Ticketing")

uploaded_files = st.sidebar.file_uploader("📁 Upload Semua File", type=["xlsx"], accept_multiple_files=True)

tiket_files, invoice_file, rekening_file, summary_files = [], None, None, []

for f in uploaded_files:
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
    for file in tiket_files:
        df = load_excel(file)
        row = df[df.apply(lambda r: r.astype(str).str.contains("TOTAL JUMLAH", regex=True).any(), axis=1)]
        if not row.empty:
            pelabuhan = next((p for p in ["merak", "bakauheni", "ketapang", "gilimanuk", "ciwandan", "panjang"] if p in file.name.lower()), "lainnya")
            b2b_data.append({"Pelabuhan": pelabuhan.capitalize(), "Pendapatan": pd.to_numeric(row.iloc[0, 4], errors="coerce")})

    invoice_df = load_excel(invoice_file)
    invoice_df["HARGA"] = pd.to_numeric(invoice_df["HARGA"], errors="coerce")
    invoice_df = invoice_df[invoice_df["STATUS"].str.lower() == "dibayar"]

    rekening_df = load_excel(rekening_file)
    rekening_detail = extract_total_rekening(rekening_df)
    rekening_total = rekening_detail["Credit"].sum()

    pengurangan_total = ""
    penambahan_dict = {}
    naik_turun_dict = {}

    match = re.search(r'(\d{4}-\d{2}-\d{2})\s*s[_\-]d\s*(\d{4}-\d{2}-\d{2})', invoice_file.name)
    if match:
        start_date, end_date = match.groups()
        start_dt, end_dt = pd.to_datetime(start_date), pd.to_datetime(end_date)
        tanggal_str = f"{start_dt.strftime('%d-%m-%Y')} s.d {end_dt.strftime('%d-%m-%Y')}"

        # Pengurangan dari summary H+1 jam 00–08
        target_date = end_dt + pd.Timedelta(days=1)
        for f in summary_files:
            if target_date.strftime('%Y-%m-%d') in f.name:
                df = load_excel(f)
                df["CETAK BOARDING PASS"] = pd.to_datetime(df["CETAK BOARDING PASS"], errors='coerce')
                df["TARIF"] = pd.to_numeric(df["TARIF"], errors="coerce")
                mask = (df["CETAK BOARDING PASS"].dt.date == target_date.date()) & \
                       (df["CETAK BOARDING PASS"].dt.time >= pd.to_datetime("00:00:00").time()) & \
                       (df["CETAK BOARDING PASS"].dt.time <= pd.to_datetime("08:00:00").time())
                pengurangan_total = df.loc[mask, "TARIF"].sum()
                break

        # Penambahan dari H jam 00–08
        for f in summary_files:
            if start_dt.strftime('%Y-%m-%d') in f.name:
                df = load_excel(f)
                df["CETAK BOARDING PASS"] = pd.to_datetime(df["CETAK BOARDING PASS"], errors='coerce')
                df["TARIF"] = pd.to_numeric(df["TARIF"], errors="coerce")
                df["ASAL"] = df["ASAL"].astype(str).str.lower()
                mask = (df["CETAK BOARDING PASS"].dt.date == start_dt.date()) & \
                       (df["CETAK BOARDING PASS"].dt.time >= pd.to_datetime("00:00:00").time()) & \
                       (df["CETAK BOARDING PASS"].dt.time <= pd.to_datetime("08:00:00").time())
                penambahan_dict = df[mask].groupby("ASAL")["TARIF"].sum().to_dict()
                break

        # Naik Turun Golongan
        invoice_df["NOMOR INVOICE"] = invoice_df.get("NOMER INVOICE", invoice_df.get("NOMOR INVOICE", "")).astype(str)
        for f in summary_files:
            if start_date in f.name and end_date in f.name:
                df = load_excel(f)
                if "NOMOR INVOICE" in df.columns and "ASAL" in df.columns:
                    df["NOMOR INVOICE"] = df["NOMOR INVOICE"].astype(str)
                    df["ASAL"] = df["ASAL"].astype(str).str.lower()
                    for pel in ["merak", "bakauheni", "ketapang", "gilimanuk", "ciwandan", "panjang"]:
                        sum_pel = df[df["ASAL"] == pel]
                        inv_pel = invoice_df[invoice_df["KEBERANGKATAN"].str.lower().str.contains(pel)]
                        ringkasan = []
                        common = set(sum_pel["NOMOR INVOICE"]).intersection(set(inv_pel["NOMOR INVOICE"]))
                        for no in common:
                            t1 = sum_pel[sum_pel["NOMOR INVOICE"] == no]["TARIF"].sum()
                            t2 = inv_pel[inv_pel["NOMOR INVOICE"] == no]["HARGA"].sum()
                            if t1 != t2:
                                ringkasan.append(f"{no}: S={int(t1)}, I={int(t2)}")
                        if ringkasan:
                            naik_turun_dict[pel] = "; ".join(ringkasan)
                break

    pelabuhan_list = ["Merak", "Bakauheni", "Ketapang", "Gilimanuk", "Ciwandan", "Panjang"]
    invoice_list = [invoice_df[invoice_df["KEBERANGKATAN"].str.lower().str.contains(p.lower())]["HARGA"].sum() for p in pelabuhan_list]
    uang_masuk_list = [rekening_total] + [0] * 5
    selisih_list = [inv - uang for inv, uang in zip(invoice_list, uang_masuk_list)]

    df = pd.DataFrame({
        "No": list(range(1, 7)),
        "Tanggal Transaksi": [tanggal_str]*6,
        "Pelabuhan Asal": pelabuhan_list,
        "Nominal Tiket Terjual": [next((b["Pendapatan"] for b in b2b_data if b["Pelabuhan"].lower() == p.lower()), 0) for p in pelabuhan_list],
        "Invoice": invoice_list,
        "Uang Masuk": uang_masuk_list,
        "Selisih": selisih_list,
        "Pengurangan": [pengurangan_total if i == 0 else "" for i in range(6)],
        "Penambahan": [penambahan_dict.get(p.lower(), "") for p in pelabuhan_list],
        "Naik Turun Golongan": [naik_turun_dict.get(p.lower(), "") for p in pelabuhan_list],
        "NET": ["" for _ in pelabuhan_list]
    })

    total_row = {
        "No": "", "Tanggal Transaksi": "", "Pelabuhan Asal": "TOTAL",
        "Nominal Tiket Terjual": df["Nominal Tiket Terjual"].sum(),
        "Invoice": df["Invoice"].sum(),
        "Uang Masuk": df["Uang Masuk"].sum(),
        "Selisih": df["Selisih"].sum(),
        "Pengurangan": "", "Penambahan": "", "Naik Turun Golongan": "", "NET": ""
    }
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    st.subheader("📄 Tabel Rekon Per Pelabuhan")
    st.dataframe(df.drop(columns=["Invoice", "Uang Masuk", "Selisih"])[df["Pelabuhan Asal"] != "TOTAL"], use_container_width=True)

    st.subheader("📄 Tabel Rekapitulasi Total")
    st.dataframe(df[df["Pelabuhan Asal"] == "TOTAL"].drop(columns=[
        "No", "Pelabuhan Asal", "Nominal Tiket Terjual", "Pengurangan", "Penambahan", "Naik Turun Golongan", "NET"
    ]), use_container_width=True)

    st.download_button("📥 Download Rekapitulasi", to_excel(df), file_name="rekapitulasi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.warning("⚠️ Silakan upload semua file yang dibutuhkan: invoice, summary, tiket, rekening.")

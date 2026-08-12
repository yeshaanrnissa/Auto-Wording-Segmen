import io
import re
import openpyxl
import pandas as pd
import streamlit as st
import math
from docx import Document


st.set_page_config(page_title="Auto Wording Segmen", layout="wide")

SEGMEN_LIST = ["Prioritas", "Payroll", "Micro", "Medium Entrepreneur"]

DEFAULT_KOLOM = {
    "Jumlah NoA": "E",
    "MTD NoA": "F",
    "% Penetrasi": "G",
    "Penjualan": "H",
    "Rata-rata Ticket Size": "J",
    "Δ MTD NoA": "N",
    "Δ MTD Penjualan": "P",
    "Δ MTD Ticket Size": "R",
}

# Format Angka
# def format_ribuan(angka):
#     if angka is None or angka == "":
#         return "0"
#     try:
#         return f"{int(float(angka)):,}".replace(",", ".")
#     except (ValueError, TypeError):
#         return "0"


# def format_desimal(angka, digit=1):
#     if angka is None or angka == "":
#         return "0"
#     try:
#         return f"{float(angka):,.{digit}f}".replace(",", "X").replace(".", ",").replace("X", ".")
#     except (ValueError, TypeError):
#         return "0"
def format_ribuan(angka):
    if angka is None or angka == "":
        return "0"
    try:
        # Pisahkan tanda minus jika ada
        string_angka = str(float(angka))
        negatif = "-" if string_angka.startswith("-") else ""
        string_angka = string_angka.lstrip("-")
        
        # Ambil bagian bulatnya saja sebelum koma desimal
        bagian_bulat = string_angka.split('.')[0]
        
        # Format ribuan untuk bagian bulat
        ribuan = f"{int(bagian_bulat):,}".replace(",", ".")
        return f"{negatif}{ribuan}"
    except (ValueError, TypeError):
        return "0"


def format_desimal(angka, digit=2):
    if angka is None or angka == "":
        return "0"
    try:
        # Menggunakan format bawaan Python (:,.2f) yang otomatis membulatkan seperti Excel
        # Ditambah trik replace standar Indonesia (Titik jadi X, Koma jadi Titik, X jadi Koma)
        return f"{float(angka):,.{digit}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "0"


def format_percentage(angka, digit=2):
    if angka is None or angka == "":
        return "0%"
    try:
        # Mengubah ke float dan dikalikan 100 untuk persentase
        nilai = float(angka) * 100
        
        # Menggunakan format bawaan (:,.f) agar otomatis membulatkan seperti Excel
        string_format = f"{nilai:,.{digit}f}"
        
        # Ubah format ribuan dan desimal sesuai standar Indonesia
        return string_format.replace(",", "X").replace(".", ",").replace("X", ".") + "%"
    except (ValueError, TypeError):
        return "0%"

# Search nama segmen

def cari_baris_segmen(sheet, keyword_mapping, max_row_scan=1000, max_col_scan=15):
    hasil = {s: None for s in SEGMEN_LIST}
    max_row = min(sheet.max_row, max_row_scan)
    max_col = min(sheet.max_column, max_col_scan)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            val = sheet.cell(row=row, column=col).value
            if not isinstance(val, str):
                continue
            val_clean = val.strip().lower()
            for segmen, keyword in keyword_mapping.items():
                if hasil[segmen] is None and val_clean == keyword.strip().lower():
                    hasil[segmen] = row
    return hasil


def cari_baris_grand_total(sheet, label_rows, grand_total_keyword="Grand Total",
                            max_row_scan=1000, max_col_scan=15):
    hasil = {s: None for s in label_rows}
    label_terurut = sorted(
        [(s, r) for s, r in label_rows.items() if r], key=lambda x: x[1]
    )
    max_row = min(sheet.max_row, max_row_scan)
    max_col = min(sheet.max_column, max_col_scan)
    kw = grand_total_keyword.strip().lower()

    for idx, (segmen, label_row) in enumerate(label_terurut):
        batas_bawah = (
            label_terurut[idx + 1][1] - 1 if idx + 1 < len(label_terurut) else max_row
        )
        for row in range(label_row + 1, batas_bawah + 1):
            for col in range(1, max_col + 1):
                val = sheet.cell(row=row, column=col).value
                if isinstance(val, str) and kw in val.strip().lower():
                    hasil[segmen] = row
                    break
            if hasil[segmen]:
                break
    return hasil


def get_row_data(sheet, row, kolom_mapping):
    if row is None or row < 1:
        return {k: None for k in kolom_mapping}
    return {
        label: sheet[f"{kolom}{row}"].value
        for label, kolom in kolom_mapping.items()
    }

# Replace paragraf di Word
def set_paragraph_text(para, new_text):
    if not para.runs:
        return
    para.runs[0].text = new_text
    for run in para.runs[1:]:
        run.text = ""

def proses_dokumen(doc, data_per_segmen):
    current_section = None
    jumlah_diubah = 0
    log = []

    for para in doc.paragraphs:
        text = para.text.strip()

        if text == "Prioritas":
            current_section = data_per_segmen.get("Prioritas")
            continue
        elif text == "Payroll":
            current_section = data_per_segmen.get("Payroll")
            continue
        elif text == "Micro":
            current_section = data_per_segmen.get("Micro")
            continue
        elif text == "Medium Entrepreneur":
            current_section = data_per_segmen.get("Medium Entrepreneur")
            continue

        if current_section is None or not para.runs:
            continue

        new_text = None
        if text.startswith("Jumlah NoA"):
            new_text = f"Jumlah NoA: {format_ribuan(current_section['Jumlah NoA'])} NoA"
        elif text.startswith("MTD NoA"):
            new_text = f"MTD NoA: {format_ribuan(current_section['MTD NoA'])} NoA"
        elif text.startswith("% Penetrasi"):
            nilai = current_section["% Penetrasi"]
            new_text = f"% Penetrasi: {format_percentage(nilai)}"
        elif text.startswith("Penjualan"):
            new_text = f"Penjualan: {format_desimal(current_section['Penjualan'])} kg"
        elif text.startswith("Rata-rata Ticket Size"):
            new_text = f"Rata-rata Ticket Size: {format_desimal(current_section['Rata-rata Ticket Size'])} gram"
        elif text.startswith("∆ MTD NoA") or text.startswith("Δ MTD NoA"):
            new_text = f"∆ MTD NoA: ({format_ribuan(current_section['Δ MTD NoA'])}) NoA"
        elif text.startswith("∆ MTD Penjualan") or text.startswith("Δ MTD Penjualan"):
            new_text = f"∆ MTD Penjualan: ({format_desimal(current_section['Δ MTD Penjualan'])}) kg"
        elif text.startswith("∆ MTD Ticket Size") or text.startswith("Δ MTD Ticket Size"):
            new_text = f"∆ MTD Ticket Size: ({format_desimal(current_section['Δ MTD Ticket Size'])}) gr"

        if new_text:
            set_paragraph_text(para, new_text)
            jumlah_diubah += 1
            log.append(f"{text[:35]!r} -> {new_text}")

    return doc, jumlah_diubah, log


# UI

import streamlit as st

# =========================
# BSI THEME - CUSTOM CSS
# =========================
st.markdown("""
    <style>
    :root {
        --bsi-teal: #00786F;
        --bsi-teal-dark: #00534D;
        --bsi-gold: #F5A623;
        --bsi-gold-dark: #D4890E;
        --bsi-bg: #F4FAF9;
    }

    /* Background utama */
    .stApp {
        background-color: var(--bsi-bg);
    }

    /* Judul & caption */
    h1 {
        color: var(--bsi-teal-dark) !important;
        font-weight: 800 !important;
    }
    h2, h3 {
        color: var(--bsi-teal) !important;
    }
    .stCaption, p, label, .st-emotion-cache-1v0mbdj {
        color: #2E2E2E;
    }

    /* Tombol utama (primary) */
    div.stButton > button[kind="primary"],
    div.stDownloadButton > button {
        background-color: var(--bsi-teal) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: 0.2s ease-in-out;
    }
    div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button:hover {
        background-color: var(--bsi-gold) !important;
        color: #1a1a1a !important;
    }

    /* File uploader */
    section[data-testid="stFileUploader"] {
        border: 2px dashed var(--bsi-teal);
        border-radius: 10px;
        background-color: #FFFFFF;
        padding: 10px;
    }

    /* Expander */
    .streamlit-expanderHeader {
        background-color: #E6F4F2;
        color: var(--bsi-teal-dark) !important;
        border-radius: 6px;
        font-weight: 600;
    }

    /* Divider */
    hr {
        border-top: 2px solid var(--bsi-gold) !important;
    }

    /* Alert / info box */
    div[data-testid="stAlert"] {
        border-left: 5px solid var(--bsi-teal);
        background-color: #E6F4F2;
    }

    /* Success box */
    div[data-baseweb="notification"] {
        background-color: #E6F4F2 !important;
        border-left: 5px solid var(--bsi-teal) !important;
    }

    /* Sidebar (jika ada) */
    section[data-testid="stSidebar"] {
        background-color: var(--bsi-teal-dark);
    }
    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    /* Dataframe header */
    .stDataFrame thead tr th {
        background-color: var(--bsi-teal) !important;
        color: white !important;
    }

    /* Number input & text input border */
    input {
        border-color: var(--bsi-teal) !important;
    }

    /* Selectbox border */
    div[data-baseweb="select"] > div {
        border-color: var(--bsi-teal) !important;
    }
    </style>
""", unsafe_allow_html=True)

# =========================
# UI
# =========================
st.title("🏦 Auto Wording Segmen Excel to Word")
st.caption(
    "Upload file Excel & template Word setiap kali laporan baru terbit — "
)

col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    excel_file = st.file_uploader("1️⃣ Upload file Excel (.xlsx)", type=["xlsx"])
with col_upload2:
    word_file = st.file_uploader("2️⃣ Upload template Word (.docx)", type=["docx"])

if excel_file:
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    sheet_names = wb.sheetnames
    default_idx = sheet_names.index("REGION") if "REGION" in sheet_names else 0
    sheet_name = st.selectbox("Pilih sheet", sheet_names, index=default_idx)
    sheet = wb[sheet_name]

    st.subheader("🔎 Deteksi Baris per Segmen")
    st.markdown("**Kata kunci pencarian label per segmen** (ubah jika istilah di Excel berbeda)")
    kw_cols = st.columns(4)
    keyword_mapping = {}
    for i, segmen in enumerate(SEGMEN_LIST):
        with kw_cols[i]:
            keyword_mapping[segmen] = st.text_input(
                f"Kata kunci - {segmen}", value=segmen, key=f"kw_{segmen}"
            )

    grand_total_keyword = st.text_input(
        "Kata kunci baris total data (baris yang diambil sebagai DATA segmen)",
        value="Grand Total",
        help="Ubah jika di Excel Anda istilahnya berbeda, misal 'Total' saja.",
    )

    cache_key = (excel_file.name, sheet_name, tuple(keyword_mapping.values()), grand_total_keyword)
    if st.session_state.get("_cache_key") != cache_key:
        label_rows = cari_baris_segmen(sheet, keyword_mapping)
        st.session_state.label_rows = label_rows
        st.session_state.auto_rows = cari_baris_grand_total(
            sheet, label_rows, grand_total_keyword
        )
        st.session_state._cache_key = cache_key

    label_rows = st.session_state.label_rows
    auto_rows = st.session_state.auto_rows

    st.markdown("**Baris data (baris 'Grand Total' yang ditemukan, bisa dikoreksi manual)**")
    baris_input = {}
    cols = st.columns(4)
    for i, segmen in enumerate(SEGMEN_LIST):
        label_row = label_rows.get(segmen)
        data_row = auto_rows.get(segmen)
        if label_row is None:
            info = "⚠️ label tidak ditemukan"
        elif data_row is None:
            info = f"label di baris {label_row}, tapi '{grand_total_keyword}' tidak ditemukan"
        else:
            info = f"label di baris {label_row} → Grand Total di baris {data_row}"
        with cols[i]:
            st.markdown(f"**{segmen}**")
            st.caption(info)
            default_val = data_row if data_row else (label_row if label_row else 1)
            baris_input[segmen] = st.number_input(
                f"Baris data {segmen}",
                min_value=1,
                value=default_val,
                key=f"row_{segmen}_{cache_key}",
            )

    with st.expander("⚙️ Pengaturan lanjutan: mapping kolom"):
        st.caption("Ubah huruf kolom jika struktur Excel berbeda dari biasanya.")
        kolom_mapping = {}
        for label, default_kolom in DEFAULT_KOLOM.items():
            kolom_mapping[label] = st.text_input(label, value=default_kolom, key=f"col_{label}")

    # Ambil data
    data_per_segmen = {
        segmen: get_row_data(sheet, baris_input[segmen], kolom_mapping)
        for segmen in SEGMEN_LIST
    }

    st.subheader("📋 Preview Data")
    df_preview = pd.DataFrame(data_per_segmen).T
    st.dataframe(df_preview, use_container_width=True)

    st.divider()

    if word_file:
        if st.button("Generate Dokumen Word", type="primary"):
            doc = Document(word_file)
            doc, jumlah_diubah, log = proses_dokumen(doc, data_per_segmen)

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.success(f"Selesai! {jumlah_diubah} paragraf berhasil diubah.")
            with st.expander("Lihat detail perubahan"):
                for line in log:
                    st.text(line)

            nama_output = re.sub(r"\.xlsx$", "", excel_file.name, flags=re.IGNORECASE)
            st.download_button(
                "⬇️ Download Hasil Word",
                data=buffer,
                file_name=f"Revised_{nama_output}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    else:
        st.info("Upload template Word untuk melanjutkan.")
else:
    st.info("Silakan upload file Excel terlebih dahulu.")

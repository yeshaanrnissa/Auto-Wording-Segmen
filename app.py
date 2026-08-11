import io
import re
import openpyxl
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="Auto Wording Segmen", layout="wide")

SEGMEN_LIST = ["Prioritas", "Payroll", "Wiraswasta", "Medium Entrepreneur"]

DEFAULT_KOLOM = {
    "Jumlah NoA": "E",
    "MTD NoA": "F",
    "% Penetrasi": "G",
    "Penjualan": "H",
    "Rata-rata Ticket Size": "J",
    "Δ MTD NoA": "O",
    "Δ MTD Penjualan": "Q",
    "Δ MTD Ticket Size": "S",
}

# ---------------------------------------------------------------------------
# Helper: format angka
# ---------------------------------------------------------------------------
def format_ribuan(angka):
    if angka is None or angka == "":
        return "0"
    try:
        return f"{int(round(float(angka))):,}".replace(",", ".")
    except (ValueError, TypeError):
        return "0"


def format_desimal(angka, digit=1):
    if angka is None or angka == "":
        return "0"
    try:
        return f"{float(angka):,.{digit}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "0"


# ---------------------------------------------------------------------------
# Helper: cari baris yang mengandung nama segmen di sheet Excel
# ---------------------------------------------------------------------------
def cari_baris_segmen(sheet, max_row_scan=1000, max_col_scan=15):
    """Scan sheet, cari baris yang selnya persis/berisi nama tiap segmen.
    Mengembalikan dict {nama_segmen: baris_ditemukan_atau_None}"""
    hasil = {s: None for s in SEGMEN_LIST}
    max_row = min(sheet.max_row, max_row_scan)
    max_col = min(sheet.max_column, max_col_scan)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            val = sheet.cell(row=row, column=col).value
            if not isinstance(val, str):
                continue
            val_clean = val.strip().lower()
            for segmen in SEGMEN_LIST:
                if hasil[segmen] is None and val_clean == segmen.lower():
                    hasil[segmen] = row
    return hasil


def get_row_data(sheet, row, kolom_mapping):
    if row is None or row < 1:
        return {k: None for k in kolom_mapping}
    return {
        label: sheet[f"{kolom}{row}"].value
        for label, kolom in kolom_mapping.items()
    }


# ---------------------------------------------------------------------------
# Helper: replace paragraf di Word
# ---------------------------------------------------------------------------
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
        elif text == "Wiraswasta":
            current_section = data_per_segmen.get("Wiraswasta")
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
            nilai_persen = nilai * 100 if nilai is not None else 0
            new_text = f"% Penetrasi: {format_desimal(nilai_persen)}%"
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


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("Auto Wording Segmen Excel to Word")
st.caption(
    "Upload file Excel & template Word setiap kali ringkasan segmen baru terbit"
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
    st.caption(
        "Sistem mencari otomatis baris berisi nama segmen. Silakan cek & koreksi "
        "nomor baris DATA (bukan baris label) di bawah ini sebelum lanjut."
    )

    if "auto_rows" not in st.session_state or st.session_state.get("_last_file") != excel_file.name:
        st.session_state.auto_rows = cari_baris_segmen(sheet)
        st.session_state._last_file = excel_file.name

    auto_rows = st.session_state.auto_rows

    baris_input = {}
    cols = st.columns(4)
    for i, segmen in enumerate(SEGMEN_LIST):
        label_row = auto_rows.get(segmen)
        info = f"label ditemukan di baris {label_row}" if label_row else "label tidak ditemukan"
        with cols[i]:
            st.markdown(f"**{segmen}**")
            st.caption(info)
            default_val = label_row if label_row else 1
            baris_input[segmen] = st.number_input(
                f"Baris data {segmen}",
                min_value=1,
                value=default_val,
                key=f"row_{segmen}",
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

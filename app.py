import streamlit as st
from docx import Document
import io
import re
import os

st.set_page_config(page_title="Jana Laporan Awal Wabak CDC", layout="wide")

st.title("📋 Penjana Laporan Awal Kejadian Wabak (CDC)")
st.write("Tampal data atau maklumat ringkas wabak di bawah untuk mengemaskini tajuk dan template laporan secara automatik.")

DEFAULT_TEMPLATE_PATH = "template.docx"

# Pilihan fail template
template_file = None
if os.path.exists(DEFAULT_TEMPLATE_PATH):
    template_file = DEFAULT_TEMPLATE_PATH
else:
    st.warning("⚠️ Fail 'template.docx' tidak dijumpai di repositori. Sila muat naik template di bawah.")
    template_file = st.file_uploader("Muat naik Fail Template (.docx)", type=["docx"])

# Input data
st.subheader("1. Tampal Data / Catatan Rawat Wabak Di Sini")
raw_data = st.text_area(
    "Data / Maklumat Rawat Wabak:",
    height=250,
    placeholder="Contoh:\nPenyakit: MEASLES\nDaerah: Kuala Selangor\nJenis Kluster/Wabak: HOUSEHOLD / MEASLES\nNama dan alamat fasiliti terlibat: Kampung Bukit Cherakah\nPemaklum: KLINIK KESIHATAN BUKIT CHERAKAH\nJumlah individu terdedah: 158 orang\nKadar Serangan: 2.53%\n..."
)

st.subheader("2. Semak & Edit Maklumat Yang Diekstrak")

def extract_field(patterns, text, default=""):
    if isinstance(patterns, str):
        patterns = [patterns]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return default

# --- Ekstraksi Automatik Berdasarkan Kata Kunci Spesifik ---

# 1 & 2. Penyakit & Jenis Kluster
raw_kluster = extract_field([r"(?:Jenis Kluster/Wabak|Jenis Kluster|Kluster|Wabak)\s*[:\-]\s*(.*)", r"(?:Penyakit)\s*[:\-]\s*(.*)"], raw_data)

# Membersihkan nama penyakit sahaja (buang HOUSEHOLD, KLUSTER, WABAK, dll)
clean_penyakit = raw_kluster
clean_penyakit = re.sub(r"(?i)\b(household|kluster|wabak)\b", "", clean_penyakit)
clean_penyakit = re.sub(r"[\/\-\:]", "", clean_penyakit).strip()

# Daerah
daerah = extract_field([r"(?:Daerah|PKD)\s*[:\-]\s*(.*)"], raw_data)

# 3. Lokaliti dari "Nama dan alamat fasiliti terlibat"
lokaliti = extract_field([r"(?:Nama dan alamat fasiliti terlibat|Lokaliti|Alamat)\s*[:\-]\s*(.*)"], raw_data)

# Pemaklum
pemaklum = extract_field([r"(?:Pemaklum|Notifikasi)\s*[:\-]\s*(.*)"], raw_data)

# 4. Bilangan Exposed dari "Jumlah individu terdedah"
exposed = extract_field([r"(?:Jumlah individu terdedah|Bilangan Exposed|Exposed)\s*[:\-]\s*(.*)"], raw_data)

# Kes Survelan
pcd = extract_field([r"(?:Jumlah Kes PCD \(Confirmed\)|Jumlah Kes PCD|PCD|Confirmed)\s*[:\-]\s*(.*)"], raw_data)
acd = extract_field([r"(?:Jumlah Kes ACD \(Suspected\)|Jumlah Kes ACD|ACD|Suspected)\s*[:\-]\s*(.*)"], raw_data)
disampel = extract_field([r"(?:Jumlah Kes Disampel|Disampel)\s*[:\-]\s*(.*)"], raw_data)
positif = extract_field([r"(?:Jumlah Kes Positif|Positif)\s*[:\-]\s*(.*)"], raw_data)
negatif = extract_field([r"(?:Jumlah Kes Negatif|Negatif)\s*[:\-]\s*(.*)"], raw_data)
pending = extract_field([r"(?:Jumlah Kes Pending|Pending)\s*[:\-]\s*(.*)"], raw_data)

# 5. Attack Rate dari "Kadar Serangan"
attack_rate = extract_field([r"(?:Kadar Serangan|Attack Rate|AR)\s*[:\-]\s*(.*)"], raw_data)

col1, col2 = st.columns(2)

with col1:
    st.markdown("### A. Maklumat Utama & Umum")
    val_penyakit = st.text_input("Nama Penyakit Sahaja (Untuk Tajuk <PENYAKIT>)", value=clean_penyakit if clean_penyakit else "MEASLES")
    val_daerah = st.text_input("Daerah (Untuk Tajuk <DAERAH>)", value=daerah)
    val_kluster = st.text_input("Jenis Kluster/Wabak", value=raw_kluster if raw_kluster else f"HOUSEHOLD / {val_penyakit}")
    val_lokaliti = st.text_area("Lokaliti (Dari 'Nama dan alamat fasiliti terlibat')", value=lokaliti)
    val_pemaklum = st.text_area("Pemaklum", value=pemaklum)

with col2:
    st.markdown("### B. Maklumat Survelan")
    val_exposed = st.text_input("Bilangan Exposed (Dari 'Jumlah individu terdedah')", value=exposed)
    val_pcd = st.text_input("Jumlah Kes PCD (Confirmed)", value=pcd)
    val_acd = st.text_input("Jumlah Kes ACD (Suspected)", value=acd)
    val_disampel = st.text_input("Jumlah Kes Disampel", value=disampel)
    val_positif = st.text_input("Jumlah Kes Positif", value=positif)
    val_negatif = st.text_input("Jumlah Kes Negatif", value=negatif)
    val_pending = st.text_input("Jumlah Kes Pending", value=pending)
    val_ar = st.text_input("Attack Rate (Dari 'Kadar Serangan')", value=attack_rate)

# Fungsi kemaskini fail Word
def replace_placeholders_and_fill_tables(source, placeholders, data_map):
    doc = Document(source)

    # 1. Gantikan <PENYAKIT> & <DAERAH> dalam perenggan/tajuk
    for p in doc.paragraphs:
        for key, val in placeholders.items():
            if key in p.text:
                p.text = p.text.replace(key, val.upper())

    # 2. Gantikan jika placeholder ada dalam jadual
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in placeholders.items():
                        if key in p.text:
                            p.text = p.text.replace(key, val.upper())

    # 3. Kemaskini kandungan jadual berasaskan label sel pertama
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) >= 2:
                key = row.cells[0].text.strip()
                if key in data_map and data_map[key] != "":
                    row.cells[1].text = data_map[key]

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

st.subheader("3. Jana & Muat Turun Laporan")

if st.button("🚀 Jana Laporan (.docx)"):
    if not template_file:
        st.error("Sila sediakan fail 'template.docx' atau muat naik fail template terlebih dahulu.")
    else:
        # Peta penggantian tajuk
        placeholders = {
            "<PENYAKIT>": val_penyakit,
            "<DAERAH>": val_daerah,
            "<penyakit>": val_penyakit,
            "<daerah>": val_daerah
        }

        # Peta pengisian jadual
        data_map = {
            "Daerah": val_daerah,
            "Jenis Kluster/Wabak": val_kluster,
            "Lokaliti": val_lokaliti,
            "Pemaklum": val_pemaklum,
            "Bilangan Exposed": val_exposed,
            "Jumlah Kes PCD (Confirmed)": val_pcd,
            "Jumlah Kes ACD (Suspected)": val_acd,
            "Jumlah Kes Disampel": val_disampel,
            "Jumlah Kes Positif": val_positif,
            "Jumlah Kes Negatif": val_negatif,
            "Jumlah Kes Pending": val_pending,
            "Attack Rate": val_ar
        }

        doc_bytes = replace_placeholders_and_fill_tables(template_file, placeholders, data_map)

        st.success("Laporan berjaya dijana!")
        st.download_button(
            label="📥 Muat Turun Laporan .docx",
            data=doc_bytes,
            file_name=f"Laporan_Awal_Wabak_{val_penyakit}_{val_daerah}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

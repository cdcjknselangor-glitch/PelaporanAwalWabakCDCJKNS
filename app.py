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
    placeholder="Tampal teks laporan / maklumat wabak di sini..."
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

# --- Ekstraksi Automatik ---

# A. Maklumat Umum
raw_kluster = extract_field([r"(?:Jenis Kluster/Wabak|Jenis Kluster|Kluster|Wabak)\s*[:\-]\s*(.*)", r"(?:Penyakit)\s*[:\-]\s*(.*)"], raw_data)
clean_penyakit = re.sub(r"(?i)\b(household|kluster|wabak)\b", "", raw_kluster)
clean_penyakit = re.sub(r"[\/\-\:]", "", clean_penyakit).strip()

daerah = extract_field([r"(?:Daerah|PKD)\s*[:\-]\s*(.*)"], raw_data)
lokaliti = extract_field([r"(?:Nama dan alamat fasiliti terlibat|Lokaliti|Alamat)\s*[:\-]\s*(.*)"], raw_data)
pemaklum = extract_field([r"(?:Pemaklum|Notifikasi)\s*[:\-]\s*(.*)"], raw_data)

# B. Maklumat Survelan
exposed = extract_field([r"(?:Jumlah individu terdedah|Bilangan Exposed|Exposed)\s*[:\-]\s*(.*)"], raw_data)
pcd = extract_field([r"(?:Jumlah Kes PCD \(Confirmed\)|Jumlah Kes PCD|PCD|Confirmed)\s*[:\-]\s*(.*)"], raw_data)
acd = extract_field([r"(?:Jumlah Kes ACD \(Suspected\)|Jumlah Kes ACD|ACD|Suspected)\s*[:\-]\s*(.*)"], raw_data)
disampel = extract_field([r"(?:Jumlah Kes Disampel|Disampel)\s*[:\-]\s*(.*)"], raw_data)
positif = extract_field([r"(?:Jumlah Kes Positif|Positif)\s*[:\-]\s*(.*)"], raw_data)
negatif = extract_field([r"(?:Jumlah Kes Negatif|Negatif)\s*[:\-]\s*(.*)"], raw_data)
pending_val = extract_field([r"(?:Jumlah Kes Pending|Pending)\s*[:\-]\s*(.*)"], raw_data, default="0")
attack_rate = extract_field([r"(?:Kadar Serangan|Attack Rate|AR)\s*[:\-]\s*(.*)"], raw_data)

# Markah Premis
markah_premis_raw = extract_field([r"(?:Markah Penilaian Risiko/Pemeriksaan Premis|Penilaian Risiko|Pemeriksaan Premis)\s*[:\-]\s*(.*)"], raw_data)
if (not markah_premis_raw or markah_premis_raw.strip() in ["", "-"] or re.search(r"deskripsi", markah_premis_raw, re.IGNORECASE)):
    markah_premis = "Tidak Berkenaan"
else:
    markah_premis = markah_premis_raw

# C. Deskripsi Kluster
onset_indeks = extract_field([r"(?:Onset Kes Indeks)\s*[:\-]\s*(.*)"], raw_data)
onset_terakhir = extract_field([r"(?:Onset Kes Terakhir)\s*[:\-]\s*(.*)"], raw_data)
umur = extract_field([r"(?:Julat Umur Kes|Julat Umur|Umur)\s*[:\-]\s*(.*)"], raw_data)
simptom = extract_field([r"(?:Simptom Utama|Simptom)\s*[:\-]\s*(.*)"], raw_data)
punca = extract_field([r"(?:Punca Jangkitan|Punca)\s*[:\-]\s*(.*)"], raw_data)

# Jenis Sampel
def extract_jenis_sampel(text):
    match = re.search(r"(?:Jenis Sampel Diambil dan Bilangan|Jenis Sampel)\s*[:\-]\s*(.*?)(?=\n[A-Z0-9\.\s]{3,}[:\-]|\Z)", text, re.IGNORECASE | re.DOTALL)
    if match:
        extracted = match.group(1).strip()
        krm_split = re.split(r"(?i)Sekiranya\s*KRM", extracted)
        extracted_clean = krm_split[0].strip()
        lines = [line.strip() for line in extracted_clean.split('\n') if line.strip()]
        return "\n".join(lines)
    return ""

jenis_sampel = extract_jenis_sampel(raw_data)

# Keputusan Sampel
text_keputusan = extract_field([r"(?:Keputusan Sampel)\s*[:\-]\s*(.*)"], raw_data)
try:
    pending_num = int(re.search(r"\d+", pending_val).group()) if re.search(r"\d+", pending_val) else 0
except:
    pending_num = 0

if text_keputusan and text_keputusan.strip().lower() not in ["", "-", "tidak berkenaan"]:
    keputusan_sampel = text_keputusan
elif pending_num > 0:
    keputusan_sampel = "Pending"
else:
    keputusan_sampel = "Tidak Berkenaan"

# D. Status Rawatan
pesakit_luar = extract_field([r"(?:Pesakit Luar)\s*[:\-]\s*(.*)"], raw_data)
wad = extract_field([r"(?:Wad)\s*[:\-]\s*(.*)"], raw_data)
icu = extract_field([r"(?:ICU)\s*[:\-]\s*(.*)"], raw_data)

# --- Peraturan 1 & 2: Maklumat Tambahan & Tindakan PKD ---
def extract_list_section(start_keyword, stop_keywords, text):
    pattern = rf"(?:{start_keyword})\s*[:\-]?\s*(.*?)(?=(?:" + "|".join(stop_keywords) + r")|\Z)"
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    items = []
    if match:
        content = match.group(1).strip()
        lines = content.split('\n')
        for line in lines:
            cleaned = re.sub(r"^\s*(?:\d+[\.\)]|[a-zA-Z][\.\)]|[\-\*])\s*", "", line).strip()
            if cleaned:
                items.append(cleaned)
    return items

list_maklumat_tambahan = extract_list_section("Maklumat Tambahan", ["Tindakan PKD", "Maklumat Pelaporan", "Dilaporkan Oleh"], raw_data)
list_tindakan_pkd = extract_list_section("Tindakan PKD", ["Maklumat Pelaporan", "Dilaporkan Oleh", "Disemak Oleh"], raw_data)

# --- Peraturan 3 & 4: Pelaporan JKN ---
tarikh_pelaporan = extract_field([r"(?:Tarikh Pelaporan)\s*[:\-]\s*(.*)"], raw_data)
masa_pelaporan = extract_field([r"(?:Masa Pelaporan)\s*[:\-]\s*(.*)"], raw_data)

# --- UI Layout ---
col1, col2 = st.columns(2)

with col1:
    st.markdown("### A. Maklumat Umum")
    val_penyakit = st.text_input("Nama Penyakit Sahaja (<PENYAKIT>)", value=clean_penyakit if clean_penyakit else "MEASLES")
    val_daerah = st.text_input("Daerah (<DAERAH>)", value=daerah)
    val_kluster = st.text_input("Jenis Kluster/Wabak", value=raw_kluster if raw_kluster else f"HOUSEHOLD / {val_penyakit}")
    val_lokaliti = st.text_area("Lokaliti", value=lokaliti)
    val_pemaklum = st.text_area("Pemaklum", value=pemaklum)

    st.markdown("### B. Maklumat Survelan")
    val_exposed = st.text_input("Bilangan Exposed", value=exposed)
    val_pcd = st.text_input("Jumlah Kes PCD (Confirmed)", value=pcd)
    val_acd = st.text_input("Jumlah Kes ACD (Suspected)", value=acd)
    val_disampel = st.text_input("Jumlah Kes Disampel", value=disampel)
    val_positif = st.text_input("Jumlah Kes Positif", value=positif)
    val_negatif = st.text_input("Jumlah Kes Negatif", value=negatif)
    val_pending = st.text_input("Jumlah Kes Pending", value=pending_val)
    val_ar = st.text_input("Attack Rate", value=attack_rate)
    val_markah_premis = st.text_input("Markah Penilaian Risiko/Premis", value=markah_premis)

    st.markdown("### G. Maklumat Pelaporan ke JKN Selangor")
    val_tarikh_pelaporan = st.text_input("Tarikh Pelaporan", value=tarikh_pelaporan)
    val_masa_pelaporan = st.text_input("Masa Pelaporan", value=masa_pelaporan)

with col2:
    st.markdown("### C. Deskripsi Kluster / Wabak")
    val_onset_indeks = st.text_input("Onset Kes Indeks", value=onset_indeks)
    val_onset_terakhir = st.text_input("Onset Kes Terakhir", value=onset_terakhir)
    val_umur = st.text_input("Julat Umur Kes", value=umur)
    val_simptom = st.text_area("Simptom Utama", value=simptom)
    val_punca = st.text_area("Punca Jangkitan", value=punca)
    val_jenis_sampel = st.text_area("Jenis Sampel", value=jenis_sampel, height=80)
    val_keputusan_sampel = st.text_input("Keputusan Sampel", value=keputusan_sampel)

    st.markdown("### D. Status Rawatan")
    val_pesakit_luar = st.text_input("Pesakit Luar", value=pesakit_luar)
    val_wad = st.text_input("Wad", value=wad)
    val_icu = st.text_input("ICU", value=icu)

    st.markdown("### E & F. Poin Maklumat & Tindakan")
    val_tambahan_text = st.text_area("E. Maklumat Tambahan (1 poin per baris)", value="\n".join(list_maklumat_tambahan), height=100)
    val_tindakan_text = st.text_area("F. Tindakan PKD (1 poin per baris)", value="\n".join(list_tindakan_pkd), height=100)

# --- Fungsi Kemaskini Fail Word & Pengurusan Jadual Dinamik ---
def fill_numbered_table(table, items):
    # Kekalkan header (row 0)
    while len(table.rows) > 1:
        # Padam baris sedia ada kecuali header
        tr = table.rows[1]._tr
        table._tbl.remove(tr)
        
    for i, item in enumerate(items, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(i)
        row_cells[1].text = item

def replace_placeholders_and_fill_tables(source, placeholders, data_map, tambahan_items, tindakan_items):
    doc = Document(source)

    # 1. Gantikan <PENYAKIT> & <DAERAH> dalam perenggan/tajuk
    for p in doc.paragraphs:
        for key, val in placeholders.items():
            if key in p.text:
                p.text = p.text.replace(key, val.upper())

    # 2. Kemaskini kandungan jadual berasaskan label
    for table in doc.tables:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        
        # Pengendalian Jadual Dinamik E & F
        if len(headers) >= 2 and headers[0] == "Bil." and headers[1] == "Maklumat":
            fill_numbered_table(table, tambahan_items)
            continue
        elif len(headers) >= 2 and headers[0] == "Bil." and headers[1] == "Tindakan":
            fill_numbered_table(table, tindakan_items)
            continue

        # Jadual Standard A, B, C, D, G
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
        placeholders = {
            "<PENYAKIT>": val_penyakit,
            "<DAERAH>": val_daerah,
            "<penyakit>": val_penyakit,
            "<daerah>": val_daerah
        }

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
            "Attack Rate": val_ar,
            "Markah Penilaian Risiko/Pemeriksaan Premis": val_markah_premis,
            "Onset Kes Indeks": val_onset_indeks,
            "Onset Kes Terakhir": val_onset_terakhir,
            "Julat Umur Kes": val_umur,
            "Simptom Utama": val_simptom,
            "Punca Jangkitan": val_punca,
            "Jenis Sampel": val_jenis_sampel,
            "Keputusan Sampel": val_keputusan_sampel,
            "Pesakit Luar": val_pesakit_luar,
            "Wad": val_wad,
            "ICU": val_icu,
            "Tarikh Pelaporan": val_tarikh_pelaporan,
            "Masa Pelaporan": val_masa_pelaporan
        }

        tambahan_list = [x.strip() for x in val_tambahan_text.split('\n') if x.strip()]
        tindakan_list = [x.strip() for x in val_tindakan_text.split('\n') if x.strip()]

        doc_bytes = replace_placeholders_and_fill_tables(template_file, placeholders, data_map, tambahan_list, tindakan_list)

        st.success("Laporan lengkap berjaya dijana!")
        st.download_button(
            label="📥 Muat Turun Laporan .docx",
            data=doc_bytes,
            file_name=f"Laporan_Awal_Wabak_{val_penyakit}_{val_daerah}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

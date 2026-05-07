import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import requests

st.set_page_config(page_title="AutoSura1372", layout="centered")

st.title("📝 Auto-Sura (Automatis Surat)")
st.write("Aplikasi ini akan mengotomasi pengisian dokumen SK Kuasa Pengguna Anggaran BPS Kota Solok.")

# --- KONFIGURASI GITHUB ---
GITHUB_RAW_URL = "https://raw.githubusercontent.com/bps1372/otomasi_sk/main/TemplateDokumen.docx"

# 1. Input Data Umum
st.subheader("Informasi Umum Dokumen")
col1, col2 = st.columns(2)

with col1:
    nomor = st.text_input("A.Nomor Dokumen", placeholder="Contoh: 042.1 TAHUN 2026")
    tanggal = st.text_input("B.Tanggal Ditetapkan", placeholder="Contoh: 15 Januari 2026")
    petugas = st.text_input("C.Menetapkan...", placeholder="Contoh: Petugas Sensus Ekonomi 2026")

with col2:
    tentang = st.text_input("D.Judul [tulis dengan huruf besar]", placeholder="Contoh: PETUGAS SENSUS EKONOMI 2026")
    pelaksanaan = st.text_input("E.Pelaksanaan Kegiatan", placeholder="Contoh: Sensus Ekonomi 2026")

# 2. Input Data Tabel (Dinamis)
st.subheader("F. Lampiran")
default_data = pd.DataFrame({
    "Nama/Jabatan": ["si ABCD"],
    "NIP/Golongan": ["Mitra Statistik"],
    "Posisi": ["PML"],
    "Honor": ["5.000.000"]
})
edited_df = st.data_editor(default_data, num_rows="dynamic", use_container_width=True)

if "doc_ready" not in st.session_state:
    st.session_state.doc_ready = False
if "doc_data" not in st.session_state:
    st.session_state.doc_data = None
if "doc_name" not in st.session_state:
    st.session_state.doc_name = ""

# --- FUNGSI CUSTOM FONT ---
def apply_bookman_font(run, size=11):
    """Menerapkan font Bookman Old Style ke dalam elemen Word"""
    run.font.name = 'Bookman Old Style'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Bookman Old Style')
    run.font.size = Pt(size)

def set_cell_text_with_font(cell, text):
    """Memasukkan teks ke sel tabel sambil mempertahankan font Bookman"""
    cell.text = text
    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        for run in p.runs:
            apply_bookman_font(run, size=11)

# 3. Tombol Eksekusi
if st.button("Proses & Buat Dokumen", type="primary"):
    
    st.info("Sedang diproses....")
    
    with st.spinner('Mengunduh template dari GitHub dan memproses dokumen...'):
        try:
            response = requests.get(GITHUB_RAW_URL)
            if response.status_code != 200:
                st.error(f"Gagal mengunduh template (Status: {response.status_code}).")
                st.session_state.doc_ready = False
            else:
                template_file = io.BytesIO(response.content)
                doc = Document(template_file)
                
                replacements = {
                    "{tentang}": tentang,
                    "{pelaksanaan}": pelaksanaan,
                    "{pelaksanan}": pelaksanaan, # Typo handling
                    "{petugas}": petugas,
                    "{tanggal}": tanggal,
                    "{nomor}": nomor
                }
                
                def replace_text_in_paragraphs(paragraphs):
                    for p in paragraphs:
                        original_text = p.text
                        changed = False
                        for key, value in replacements.items():
                            if key in original_text:
                                original_text = original_text.replace(key, value)
                                changed = True
                        if changed:
                            p.text = original_text
                            for run in p.runs:
                                apply_bookman_font(run, size=12)

                replace_text_in_paragraphs(doc.paragraphs)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if "{nama}" not in cell.text.lower():
                                replace_text_in_paragraphs(cell.paragraphs)

                # --- Logika Tabel Lampiran ---
                target_table = None
                template_row_idx = -1
                for table in doc.tables:
                    for i, row in enumerate(table.rows):
                        for cell in row.cells:
                            if "{nama}" in cell.text.lower():
                                target_table = table
                                template_row_idx = i
                                break
                        if target_table: break
                    if target_table: break
                
                if target_table and template_row_idx != -1:
                    # Ambil elemen XML dari baris template sebagai titik awal
                    current_tr = target_table.rows[template_row_idx]._tr
                    
                    for index, row_data in edited_df.iterrows():
                        if index == 0:
                            # Data pertama menimpa baris template {nama}
                            target_row = target_table.rows[template_row_idx]
                        else:
                            # 1. BUAT BARIS KOSONG (SPACER) SEBAGAI JARAK ANTAR DATA (1 Spasi)
                            spacer_row = target_table.add_row()
                            for cell in spacer_row.cells:
                                p = cell.paragraphs[0]
                                p.paragraph_format.space_before = Pt(0)
                                p.paragraph_format.space_after = Pt(0)
                                p.paragraph_format.line_spacing = 1.0 
                                run = p.add_run("")
                                apply_bookman_font(run, size=11)
                            
                            # Sisipkan baris kosong tepat di bawah baris terakhir
                            current_tr.addnext(spacer_row._tr)
                            current_tr = spacer_row._tr
                            
                            # 2. BUAT BARIS UNTUK DATA BARU
                            new_row = target_table.add_row()
                            
                            # Sisipkan baris data tepat di bawah baris kosong (spacer)
                            current_tr.addnext(new_row._tr)
                            current_tr = new_row._tr 
                            
                            target_row = new_row
                        
                        # Isi data ke dalam baris
                        set_cell_text_with_font(target_row.cells[0], f"{index + 1}.")
                        set_cell_text_with_font(target_row.cells[1], str(row_data["Nama/Jabatan"]))
                        set_cell_text_with_font(target_row.cells[2], str(row_data["NIP/Golongan"]))
                        set_cell_text_with_font(target_row.cells[3], str(row_data["Posisi"]))
                        set_cell_text_with_font(target_row.cells[4], f"Rp{row_data['Honor']}")

                    # 3. TAMBAHKAN JARAK KHUSUS (1.5 SPASI) SETELAH SEMUA DATA SELESAI
                    # Ini akan memberi jarak antara data paling terakhir dengan garis ganda penutup
                    final_spacer_row = target_table.add_row()
                    for cell in final_spacer_row.cells:
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.line_spacing = 1.5 # Jarak 1.5 khusus bagian paling bawah
                        run = p.add_run("")
                        apply_bookman_font(run, size=11)
                    
                    # Sisipkan tepat setelah data terakhir
                    current_tr.addnext(final_spacer_row._tr)

                else:
                    st.warning("⚠️ Peringatan: Teks {nama} tidak ditemukan dalam tabel.")

                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                st.session_state.doc_data = doc_io.getvalue()
                st.session_state.doc_name = f"SK_{tentang.replace(' ', '_')}.docx"
                st.session_state.doc_ready = True
                
        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {e}")
            st.session_state.doc_ready = False

# 4. Menampilkan Tombol Unduh
if st.session_state.doc_ready:
    st.success("✅ Dokumen berhasil diproses")
    st.download_button(
        label="⬇️ Unduh Dokumen Hasil",
        data=st.session_state.doc_data,
        file_name=st.session_state.doc_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

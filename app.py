import streamlit as st
from openai import OpenAI
import json
import re
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Generator Modul Ajar AI", page_icon="📚", layout="wide")

st.title("Generator Modul Ajar Praktis (Format Word DOCX) 🤖📚")
st.markdown("Masukkan Mapel dan Bab, AI akan membuatkan Modul Ajar **Kurikulum Berbasis Cinta - Deep Learning** langsung ke dalam tabel Word (DOCX) yang rapi.")

# --- API KEY ---
try:
    api_key = st.secrets["NVIDIA_API_KEY"]
except KeyError:
    st.error("⚠️ NVIDIA_API_KEY tidak ditemukan di secrets. Pastikan Anda sudah mengaturnya di `.streamlit/secrets.toml`.")
    st.stop()

client = OpenAI(
    base_url="https://integrate.api.nvidia.com/v1",
    api_key=api_key
)

# --- FUNGSI PEMBUAT DOCX DENGAN TABEL RAPI ---
def generate_docx(meta, data):
    doc = Document()
    
    # Judul
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("MODUL AJAR\nKURIKULUM BERBASIS CINTA – PENDEKATAN DEEP LEARNING")
    run.bold = True
    run.font.size = Pt(14)
    
    def add_table_data(title, rows_data):
        if title:
            doc.add_heading(title, level=2)
        table = doc.add_table(rows=len(rows_data), cols=len(rows_data[0]))
        table.style = 'Table Grid'
        for i, row_data in enumerate(rows_data):
            for j, cell_data in enumerate(row_data):
                table.rows[i].cells[j].text = str(cell_data)
        doc.add_paragraph()
        return table

    # 1. Tabel Identitas
    doc.add_heading('IDENTITAS MODUL AJAR', level=2)
    t_id = doc.add_table(rows=4, cols=4)
    t_id.style = 'Table Grid'
    
    r0 = t_id.rows[0].cells
    r0[0].text, r0[1].text, r0[2].text, r0[3].text = "Mata Pelajaran", meta['mapel'], "Kelas / Fase", meta['kelas']
    
    r1 = t_id.rows[1].cells
    r1[0].text, r1[1].text, r1[2].text, r1[3].text = "Semester", meta['semester'], "Alokasi Waktu", meta['waktu']
    
    r2 = t_id.rows[2].cells
    r2[0].text, r2[1].text = "Bab / Topik", meta['topik']
    r2[1].merge(r2[3]) # Menggabungkan kolom bab/topik agar panjang
    
    r3 = t_id.rows[3].cells
    r3[0].text, r3[1].text, r3[2].text, r3[3].text = "Penyusun", meta['guru'], "Sekolah", meta['sekolah']
    doc.add_paragraph()

    # 2. Tabel Identifikasi Peserta Didik
    add_table_data('A. IDENTIFIKASI PESERTA DIDIK', [
        ["Pengetahuan Awal", data.get("pengetahuan_awal", "")],
        ["Minat Belajar", data.get("minat_belajar", "")],
        ["Latar Belakang", data.get("latar_belakang", "")],
        ["Kebutuhan Belajar", data.get("kebutuhan_belajar", "")],
        ["Dimensi Profil Kelulusan", data.get("dimensi_profil", "")],
        ["Topik Panca Cinta", data.get("panca_cinta", "")]
    ])

    # 3. Tabel Desain Pembelajaran
    add_table_data('B. DESAIN PEMBELAJARAN', [
        ["Capaian Pembelajaran", data.get("cp", "")],
        ["Tujuan Pembelajaran", data.get("tp", "")],
        ["Lintas Disiplin Ilmu", data.get("lintas_disiplin", "")],
        ["Topik Pembelajaran", data.get("topik_pembelajaran", "")],
        ["Praktik Pedagogi", "Model: PBL (Problem Based Learning)\nMetode: Ceramah Interaktif, Diskusi Kelompok, Tanya Jawab\nPrinsip DL: Mindful, Meaningful, Joyful"]
    ])

    # 4. Tabel Pengalaman Belajar
    doc.add_heading('C. PENGALAMAN BELAJAR', level=2)
    doc.add_paragraph(f"Materi: {data.get('pertemuan_1_materi', '')}")
    add_table_data('', [
        ["FASE KEGIATAN", "AKTIVITAS PEMBELAJARAN", "PRINSIP DL"],
        ["PEMBUKAAN", data.get("pertemuan_1_pembukaan", ""), "MEANINGFUL (Bermakna)"],
        ["KEGIATAN INTI\n(Sintak PBL)", data.get("pertemuan_1_inti", ""), "MINDFUL & JOYFUL"],
        ["PENUTUP", data.get("pertemuan_1_penutup", ""), "MINDFUL (Berkesadaran)"]
    ])

    # 5. Tabel Asesmen
    add_table_data('D. PENILAIAN / ASESMEN', [
        ["Asesmen Awal (Diagnostik)", data.get("asesmen_awal", "")],
        ["Asesmen Formatif", data.get("asesmen_formatif", "")],
        ["Asesmen Sumatif", data.get("asesmen_sumatif", "")]
    ])
    
    doc.add_page_break()
    
    # 6. Lampiran Materi & LKPD
    doc.add_heading('LAMPIRAN MATERI AJAR', level=2)
    doc.add_paragraph(data.get("materi_ajar", ""))
    
    doc.add_heading('LEMBAR KERJA PESERTA DIDIK (LKPD)', level=2)
    doc.add_paragraph(data.get("lkpd", ""))

    # 7. Tanda Tangan
    doc.add_paragraph("\n\n")
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.cell(0,0).text = "Mengetahui,\nKepala Sekolah\n\n\n\n\n( ________________________ )"
    sig_table.cell(0,1).text = f"Guru Mata Pelajaran\n\n\n\n\n( {meta['guru']} )"
    
    # Rata tengah teks tanda tangan
    for row in sig_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    return doc

# --- FORM INPUT GURU ---
st.markdown("### 📝 Silakan Isi Data Modul")
with st.form("form_modul"):
    col1, col2 = st.columns(2)
    with col1:
        mata_pelajaran = st.text_input("Mata Pelajaran", placeholder="Contoh: Ilmu Pengetahuan Alam")
        kelas_fase = st.text_input("Kelas / Fase", placeholder="Contoh: VI / Fase C")
        topik = st.text_input("Bab / Topik Pembelajaran", placeholder="Contoh: Tata Surya")
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        nama_guru = st.text_input("Nama Guru Penyusun", placeholder="Contoh: Budi Santoso, S.Pd")
        sekolah = st.text_input("Nama Sekolah", placeholder="Contoh: SD Negeri 1 Bogor")
        alokasi_waktu = st.text_input("Alokasi Waktu", placeholder="Contoh: 4 JP (2 Pertemuan)")

    submit_button = st.form_submit_button("✨ Buat Modul Ajar Sekarang")

# --- PROSES AI & GENERATE DOCX ---
if submit_button:
    if not mata_pelajaran or not topik:
        st.warning("⚠️ Mohon isi minimal Mata Pelajaran dan Bab/Topik.")
    else:
        with st.spinner("Meracik modul dan menggambar tabel Word... Ini membutuhkan waktu beberapa detik."):
            try:
                # Prompt meminta AI mengembalikan format JSON yang ketat
                system_prompt = f"""
                Anda adalah pembuat Modul Ajar ahli. Buatlah modul ajar yang SANGAT DETAIL berdasarkan "Kurikulum Berbasis Cinta - Pendekatan Deep Learning".
                Topik: {topik}, Mapel: {mata_pelajaran}.
                
                WAJIB berikan respons HANYA dalam bentuk format JSON yang valid, tanpa teks markdown di luarnya.
                Gunakan bullet (titik/tanda strip) untuk merapikan teks di dalam nilai string.
                
                Gunakan keys JSON berikut persis seperti ini:
                {{
                  "pengetahuan_awal": "...",
                  "minat_belajar": "...",
                  "latar_belakang": "...",
                  "kebutuhan_belajar": "...",
                  "dimensi_profil": "...",
                  "panca_cinta": "Kaitan 5 panca cinta dengan materi ini...",
                  "cp": "Capaian pembelajaran...",
                  "tp": "Tujuan pembelajaran (gunakan tingkatan kognitif)...",
                  "lintas_disiplin": "Kaitan dengan mapel lain...",
                  "topik_pembelajaran": "Sub-topik yang akan dibahas...",
                  "pertemuan_1_materi": "Judul spesifik materi pertemuan...",
                  "pertemuan_1_pembukaan": "Langkah pembukaan yang Meaningful...",
                  "pertemuan_1_inti": "Langkah inti PBL (Orientasi, Organisasi, Penyelidikan, Presentasi)...",
                  "pertemuan_1_penutup": "Langkah penutup...",
                  "asesmen_awal": "Pertanyaan diagnostik...",
                  "asesmen_formatif": "Cara asesmen formatif...",
                  "asesmen_sumatif": "Berikan 3 contoh soal evaluasi HOTS...",
                  "materi_ajar": "Ringkasan materi yang padat dan jelas...",
                  "lkpd": "Instruksi dan pertanyaan lembar kerja siswa..."
                }}
                """

                # Mengirim request ke NVIDIA AI
                completion = client.chat.completions.create(
                    model="nvidia/nemotron-3-ultra-550b-a55b",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": "Berikan respons HANYA dalam format JSON murni. Jangan ada teks pembuka atau penutup sama sekali."}
                    ],
                    temperature=0.2, # Suhu rendah agar AI disiplin pada format
                    max_tokens=4000 
                )

                hasil_ai = completion.choices[0].message.content
                
                # --- PEMBERSIHAN JSON ---
                teks_json = hasil_ai.strip()
                
                # Buang tag markdown jika AI masih menyertakannya
                if teks_json.startswith("```json"):
                    teks_json = teks_json[7:]
                elif teks_json.startswith("```"):
                    teks_json = teks_json[3:]
                    
                if teks_json.endswith("```"):
                    teks_json = teks_json[:-3]
                    
                teks_json = teks_json.strip()
                
                # Parsing string JSON menjadi dictionary Python
                try:
                    data_json = json.loads(teks_json) 
                except json.JSONDecodeError as e:
                    st.error("⚠️ AI gagal merangkai format JSON dengan sempurna. Silakan klik tombol 'Buat Modul' sekali lagi.")
                    with st.expander("Lihat Output Mentah AI (Untuk Debugging)"):
                        st.error(f"Detail kode error: {e}")
                        st.text_area("Apa yang AI katakan:", hasil_ai, height=300)
                    st.stop()

                # --- LANJUT KE PEMBUATAN DOCX ---
                meta_data = {
                    "mapel": mata_pelajaran,
                    "kelas": kelas_fase,
                    "semester": semester,
                    "waktu": alokasi_waktu,
                    "topik": topik,
                    "guru": nama_guru if nama_guru else "_______________",
                    "sekolah": sekolah if sekolah else "_______________"
                }

                # Generate File Word menggunakan Python-Docx
                doc = generate_docx(meta_data, data_json)
                
                # Simpan ke memori (bytes) untuk tombol download
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                st.success("✅ Modul Ajar berhasil dibuat dan tabel sudah dirapikan!")
                
                # Tombol Download DOCX
                st.download_button(
                    label="📥 Download Modul (Format Word / DOCX)",
                    data=doc_io,
                    file_name=f"Modul_{mata_pelajaran.replace(' ', '_')}_{topik.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Terjadi kesalahan pada sistem: {e}")

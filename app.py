import streamlit as st
from openai import OpenAI
import json
import re
import io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Generator Modul Ajar AI", page_icon="📚", layout="wide")

st.title("Generator Modul Ajar Premium (DOCX Berwarna) 🤖🎨")
st.markdown("Menghasilkan Modul Ajar dengan format tabel **berwarna dan rapi** persis seperti template asli.")

# --- API KEY ---
try:
    api_key = st.secrets["NVIDIA_API_KEY"]
except KeyError:
    st.error("⚠️ NVIDIA_API_KEY tidak ditemukan di secrets. Pastikan sudah ada di .streamlit/secrets.toml")
    st.stop()

client = OpenAI(
    base_url="https://integrate.api.nvidia.com/v1",
    api_key=api_key
)

# --- FUNGSI HELPER UNTUK WARNA & TABEL ---
def set_cell_bg(cell, color_hex="D9E2F3"):
    """Fungsi untuk memberi warna background (shading) pada sel tabel Word."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
    tcPr.append(shd)

def style_header_cell(cell, text, bg_color="2F5496", text_color=RGBColor(255, 255, 255)):
    """Memformat sel sebagai Header (Warna gelap, Teks putih tebal)."""
    set_cell_bg(cell, bg_color)
    cell.text = text
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.bold = True
            run.font.color.rgb = text_color

def insert_bullet_points(cell, text_data):
    """Mengubah teks dengan baris baru menjadi bullet points di dalam sel tabel."""
    lines = str(text_data).split('\n')
    for i, line in enumerate(lines):
        line = line.strip().strip('-').strip('•').strip()
        if line:
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.text = line
            p.style = 'List Bullet'

# --- FUNGSI PEMBUAT DOCX ---
def generate_docx(meta, data):
    doc = Document()
    
    # 1. Judul Utama
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("MODUL AJAR\n")
    r1.bold = True
    r1.font.size = Pt(16)
    r2 = p.add_run("KURIKULUM BERBASIS CINTA – PENDEKATAN DEEP LEARNING\nBerdasarkan Capaian Pembelajaran Terbaru 2025")
    r2.bold = True
    r2.font.size = Pt(12)

    # 2. Tabel Identitas
    doc.add_heading('IDENTITAS MODUL AJAR', level=2)
    t_id = doc.add_table(rows=5, cols=4)
    t_id.style = 'Table Grid'
    
    # Rata kiri kolom ganjil diberi warna biru muda
    id_data = [
        ("Mata Pelajaran", meta['mapel'], "Kelas / Fase", meta['kelas']),
        ("Semester", meta['semester'], "Alokasi Waktu", meta['waktu']),
        ("Bab / Topik", meta['topik'], "", ""),
        ("Model Pembelajaran", "PBL (Problem Based Learning)", "Metode Pembelajaran", data.get("metode", "Ceramah, Diskusi")),
        ("Penyusun", meta['guru'], "Sekolah", meta['sekolah'])
    ]
    
    for i, row_data in enumerate(id_data):
        cells = t_id.rows[i].cells
        cells[0].text, cells[1].text = row_data[0], row_data[1]
        set_cell_bg(cells[0], "D9E2F3") # Warna kolom kiri
        
        if row_data[2]: # Jika ada data kolom kanan
            cells[2].text, cells[3].text = row_data[2], row_data[3]
            set_cell_bg(cells[2], "D9E2F3")
        else:
            cells[1].merge(cells[3]) # Gabung sel jika kosong (untuk Bab/Topik)

    doc.add_paragraph()

    # 3. Tabel Identifikasi Peserta Didik
    doc.add_heading('A. IDENTIFIKASI PESERTA DIDIK', level=2)
    t_peserta = doc.add_table(rows=6, cols=2)
    t_peserta.style = 'Table Grid'
    t_peserta.columns[0].width = Pt(150) # Kolom kiri lebih kecil
    
    peserta_fields = [
        ("Pengetahuan Awal", data.get("pengetahuan_awal", "")),
        ("Minat Belajar", data.get("minat_belajar", "")),
        ("Latar Belakang", data.get("latar_belakang", "")),
        ("Kebutuhan Belajar", data.get("kebutuhan_belajar", "")),
        ("Dimensi Profil Kelulusan", data.get("dimensi_profil", "")),
        ("Topik Panca Cinta", data.get("panca_cinta", ""))
    ]
    
    for i, (k, v) in enumerate(peserta_fields):
        t_peserta.cell(i, 0).text = k
        set_cell_bg(t_peserta.cell(i, 0), "D9E2F3")
        t_peserta.cell(i, 0).paragraphs[0].runs[0].font.bold = True
        insert_bullet_points(t_peserta.cell(i, 1), v)

    doc.add_paragraph()

    # 4. Tabel Desain Pembelajaran
    doc.add_heading('B. DESAIN PEMBELAJARAN', level=2)
    t_desain = doc.add_table(rows=8, cols=2)
    t_desain.style = 'Table Grid'
    
    desain_fields = [
        ("Capaian Pembelajaran (CP)", data.get("cp", "")),
        ("Tujuan Pembelajaran (TP)", data.get("tp", "")),
        ("Lintas Disiplin Ilmu", data.get("lintas_disiplin", "")),
        ("Topik Pembelajaran", data.get("sub_topik", "")),
        ("Praktik Pedagogi", "• Model: PBL\n• Prinsip: Mindful, Meaningful, Joyful"),
        ("Lingkungan Belajar", data.get("lingkungan_belajar", "")),
        ("Kemitraan Pembelajaran", data.get("kemitraan", "")),
        ("Pemanfaatan Digital", data.get("digital", ""))
    ]
    
    for i, (k, v) in enumerate(desain_fields):
        t_desain.cell(i, 0).text = k
        set_cell_bg(t_desain.cell(i, 0), "D9E2F3")
        t_desain.cell(i, 0).paragraphs[0].runs[0].font.bold = True
        insert_bullet_points(t_desain.cell(i, 1), v)
        
    doc.add_paragraph()

    # 5. Tabel Pengalaman Belajar
    doc.add_heading('C. PENGALAMAN BELAJAR', level=2)
    p_pengalaman = doc.add_paragraph()
    p_pengalaman.add_run(f"Materi: {meta['topik']} | Durasi: {meta['waktu']} | Model: PBL").bold = True
    
    t_pengalaman = doc.add_table(rows=4, cols=3)
    t_pengalaman.style = 'Table Grid'
    
    # Header berwarna Biru Gelap
    headers = ["FASE KEGIATAN", "AKTIVITAS PEMBELAJARAN", "PRINSIP DL"]
    for i, h in enumerate(headers):
        style_header_cell(t_pengalaman.cell(0, i), h)
        
    # Baris Isi
    fase_data = [
        ("PEMBUKAAN", data.get("kegiatan_pembukaan", ""), "MEANINGFUL\n(Bermakna)"),
        ("MEMAHAMI & MENGAPLIKASIKAN\n(Langkah 1-4 PBL)", data.get("kegiatan_inti", ""), "MINDFUL &\nJOYFUL"),
        ("MEREFLEKSIKAN & PENUTUP", data.get("kegiatan_penutup", ""), "MEANINGFUL")
    ]
    
    for i, (fase, aktivitas, prinsip) in enumerate(fase_data, start=1):
        t_pengalaman.cell(i, 0).text = fase
        set_cell_bg(t_pengalaman.cell(i, 0), "F2F2F2") # Abu-abu terang
        insert_bullet_points(t_pengalaman.cell(i, 1), aktivitas)
        t_pengalaman.cell(i, 2).text = prinsip

    doc.add_paragraph()

    # 6. Lampiran (Halaman Baru)
    doc.add_page_break()
    doc.add_heading('LAMPIRAN - MATERI & LKPD', level=1)
    
    doc.add_heading('A. Ringkasan Materi', level=2)
    doc.add_paragraph(data.get("materi_ajar", ""))
    
    doc.add_heading('B. Lembar Kerja Peserta Didik (LKPD)', level=2)
    doc.add_paragraph(data.get("lkpd", ""))
    
    doc.add_heading('C. Asesmen Sumatif (Soal HOTS)', level=2)
    insert_bullet_points(doc.add_paragraph(), data.get("soal_hots", ""))

    # 7. Tanda Tangan
    doc.add_paragraph("\n\n")
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.cell(0,0).text = "Mengetahui,\nKepala Sekolah\n\n\n\n\n( ________________________ )"
    sig_table.cell(0,1).text = f"Guru Mata Pelajaran\n\n\n\n\n( {meta['guru']} )"
    
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
        mata_pelajaran = st.text_input("Mata Pelajaran", placeholder="Contoh: Bahasa Indonesia")
        kelas_fase = st.text_input("Kelas / Fase", placeholder="Contoh: VI / Fase C")
        topik = st.text_input("Bab / Topik Pembelajaran", placeholder="Contoh: Anak-Anak yang Mengubah Dunia")
        semester = st.selectbox("Semester", ["1 (Satu)", "2 (Dua)"])
    with col2:
        nama_guru = st.text_input("Nama Guru Penyusun", placeholder="Contoh: Mamik Muhapatin, S.Pd")
        sekolah = st.text_input("Nama Sekolah", placeholder="Contoh: MI. Miftahussalam")
        alokasi_waktu = st.text_input("Alokasi Waktu", placeholder="Contoh: 24 JP (6 Pertemuan)")

    submit_button = st.form_submit_button("✨ Buat Modul Berwarna Sekarang")

# --- PROSES AI & GENERATE DOCX ---
if submit_button:
    if not mata_pelajaran or not topik:
        st.warning("⚠️ Mohon isi minimal Mata Pelajaran dan Bab/Topik.")
    else:
        with st.spinner("Meracik modul dan mewarnai tabel Word... Mohon tunggu sekitar 15-30 detik."):
            try:
                system_prompt = f"""
                Anda adalah pembuat Modul Ajar ahli dengan menggunakan CP sesuai Keputusan Mentri Agama 1503 tahun 2025. Buatlah modul ajar SANGAT MENDALAM berdasarkan "Kurikulum Berbasis Cinta - Pendekatan Deep Learning".
                Topik: {topik}, Mapel: {mata_pelajaran}, Fase: {kelas_fase}.
                
                
                WAJIB berikan respons HANYA format JSON valid. Jangan berikan teks lain!
                Pisahkan setiap poin dengan karakter "\\n-" (newline dan strip) agar bisa dijadikan bullet point.
                
                Gunakan keys JSON berikut:
                {{
                  "metode": "Ceramah Interaktif, Diskusi Kelompok, Proyek",
                  "pengetahuan_awal": "Pengetahuan awal siswa terkait {topik}...",
                  "minat_belajar": "Minat siswa kelas ini...",
                  "latar_belakang": "Latar belakang era digital...",
                  "kebutuhan_belajar": "Visual: ...\\n- Audio: ...\\n- Kinestetik: ...",
                  "dimensi_profil": "1. Bernalar Kritis: ...\\n- 2. Kreatif: ...\\n- 3. Gotong Royong: ...",
                  "panca_cinta": "1. Cinta Allah: ...\\n- 2. Cinta Sesama: ...\\n- 3. Cinta Ilmu: ...\\n- 4. Cinta Lingkungan: ...\\n- 5. Cinta Tanah Air: ...",
                  "cp": "Tuliskan CP yang relevan...",
                  "tp": "TP 1 (Pemahaman Dasar): ...\\n- TP 2 (Berpikir Kritis): ...",
                  "lintas_disiplin": "Kaitan dengan mapel lain (misal IPS/IPA)...",
                  "sub_topik": "Rincian sub bab yang dibahas...",
                  "lingkungan_belajar": "Setting kelas...",
                  "kemitraan": "Guru mapel lain, orang tua...",
                  "digital": "Canva, Quizizz, Youtube...",
                  "kegiatan_pembukaan": "Guru mengucapkan salam...\\n- Pertanyaan pemantik...\\n- Motivasi...",
                  "kegiatan_inti": "Langkah 1 (Orientasi Masalah): ...\\n- Langkah 2 (Organisasi): ...\\n- Langkah 3 (Penyelidikan): ...",
                  "kegiatan_penutup": "Kesimpulan...\\n- Refleksi...\\n- Doa...",
                  "materi_ajar": "Tuliskan ringkasan materi {topik} yang padat...",
                  "lkpd": "Instruksi LKPD:\\n- Tugas 1...\\n- Tugas 2...",
                  "soal_hots": "1. Soal analisis...\\n- 2. Soal evaluasi...\\n- 3. Soal kreasi..."
                }}
                """

                # Mengirim request ke NVIDIA AI
                completion = client.chat.completions.create(
                    model="nvidia/nemotron-3-ultra-550b-a55b",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": "Keluarkan JSON murni tanpa ada tag markdown."}
                    ],
                    temperature=0.1, 
                    max_tokens=4000 
                )

                hasil_ai = completion.choices[0].message.content
                
                # Pembersihan JSON dengan cara yang aman untuk disalin-tempel
                teks_json = hasil_ai.strip()
                ticks = "`" * 3
                
                if teks_json.startswith(ticks + "json"):
                    teks_json = teks_json[7:]
                elif teks_json.startswith(ticks):
                    teks_json = teks_json[3:]
                    
                if teks_json.endswith(ticks):
                    teks_json = teks_json[:-3]
                    
                teks_json = teks_json.strip()
                
                # Parsing string JSON menjadi dictionary Python
                try:
                    data_json = json.loads(teks_json) 
                except json.JSONDecodeError as e:
                    st.error("⚠️ AI gagal merangkai format JSON dengan sempurna.")
                    with st.expander("Lihat Output Mentah AI (Untuk Debugging)"):
                        st.error(f"Detail kode error: {e}")
                        st.text_area("Apa yang AI katakan:", hasil_ai, height=300)
                    st.stop()

                # Metadata untuk identitas
                meta_data = {
                    "mapel": mata_pelajaran,
                    "kelas": kelas_fase,
                    "semester": semester,
                    "waktu": alokasi_waktu,
                    "topik": topik,
                    "guru": nama_guru if nama_guru else "_______________",
                    "sekolah": sekolah if sekolah else "_______________"
                }

                # Generate File Word berwarna
                doc = generate_docx(meta_data, data_json)
                
                # Simpan ke memori untuk tombol download
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                st.success("✅ Modul Ajar Berwarna berhasil dibuat!")
                
                # Tombol Download DOCX
                st.download_button(
                    label="📥 Download Modul (Format Word Berwarna)",
                    data=doc_io,
                    file_name=f"Modul_{mata_pelajaran.replace(' ', '_')}_{topik.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Terjadi kesalahan pada sistem: {e}")

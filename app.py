import streamlit as st
from openai import OpenAI

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Generator Modul Ajar AI", page_icon="📚", layout="wide")

st.title("Generator Modul Ajar Praktis 🤖📚")
st.markdown("Cukup masukkan Mapel dan Bab, AI akan membuatkan Modul Ajar **Kurikulum Berbasis Cinta - Deep Learning** yang lengkap beserta LKPD dan Asesmen.")

# Mengambil API Key dari Streamlit Secrets (NVIDIA Nemotron)
try:
    api_key = st.secrets["NVIDIA_API_KEY"]
except KeyError:
    st.error("⚠️ NVIDIA_API_KEY tidak ditemukan di secrets.")
    st.stop()

# Inisialisasi Client OpenAI untuk NVIDIA API
client = OpenAI(
    base_url="https://integrate.api.nvidia.com/v1",
    api_key=api_key
)

# === ANTARMUKA SEDERHANA UNTUK GURU ===
st.markdown("### 📝 Silakan Isi Data Modul")
with st.form("form_modul"):
    col1, col2 = st.columns(2)
    
    with col1:
        mata_pelajaran = st.text_input("Mata Pelajaran", placeholder="Contoh: Matematika")
        kelas_fase = st.text_input("Kelas / Fase", placeholder="Contoh: VI / Fase C")
        topik = st.text_input("Bab / Topik Pembelajaran", placeholder="Contoh: Pecahan dan Desimal")
        
    with col2:
        nama_guru = st.text_input("Nama Guru Penyusun", placeholder="Contoh: Mamik Muhapatin, S.Pd")
        sekolah = st.text_input("Nama Sekolah", placeholder="Contoh: MI. Miftahussalam")
        alokasi_waktu = st.text_input("Alokasi Waktu", placeholder="Contoh: 2 Pertemuan (4 JP)")

    submit_button = st.form_submit_button("✨ Buat Modul Ajar Sekarang")

# === PROSES GENERASI OLEH AI ===
if submit_button:
    if not mata_pelajaran or not topik:
        st.warning("⚠️ Mohon isi minimal Mata Pelajaran dan Bab/Topik terlebih dahulu.")
    else:
        with st.spinner("AI sedang merancang modul, LKPD, dan Asesmen... Mohon tunggu sebentar."):
            try:
                # Prompt yang didesain khusus agar meniru format PDF referensi Anda
                system_prompt = f"""
                Anda adalah pembuat Modul Ajar ahli. Buatlah modul ajar yang SANGAT DETAIL berdasarkan "Kurikulum Berbasis Cinta - Pendekatan Deep Learning 2025".
                Semua konten (materi, aktivitas, asesmen) harus HANYA tentang Mata Pelajaran: {mata_pelajaran} dan Topik: {topik}.
                
                Gunakan format Markdown dan ikuti struktur dokumen persis seperti ini:
                
                # MODUL AJAR {mata_pelajaran.upper()}
                **Kurikulum Berbasis Cinta - Deep Learning 2025**
                
                ## IDENTITAS MODUL AJAR
                - Mata Pelajaran: {mata_pelajaran}
                - Kelas / Fase: {kelas_fase}
                - Bab / Topik: {topik}
                - Alokasi Waktu: {alokasi_waktu}
                - Penyusun: {nama_guru}
                - Sekolah: {sekolah}
                - Model Pembelajaran: PBL (Problem Based Learning)
                
                ## A. IDENTIFIKASI PESERTA DIDIK
                - Pengetahuan Awal
                - Minat Belajar
                - Latar Belakang
                - Kebutuhan Belajar
                - Dimensi Profil Kelulusan (DPL)
                - Topik Panca Cinta (Cinta Allah, Diri, Ilmu, Lingkungan, Tanah Air - kaitkan dengan {topik})
                
                ## B. DESAIN PEMBELAJARAN
                - Capaian Pembelajaran (CP)
                - Tujuan Pembelajaran (TP)
                - Lintas Disiplin Ilmu
                - Praktik Pedagogi (PBL)
                
                ## C. PENGALAMAN BELAJAR
                Buat aktivitas pertemuan dengan tabel Markdown (Fase, Kegiatan, Waktu, Prinsip Deep Learning).
                Bagi menjadi Fase: PEMBUKAAN (Meaningful), INTI (Sintak PBL: Mindful, Joyful, Meaningful), PENUTUP.
                Pastikan materi pembelajarannya fokus pada {topik}.
                
                ## D. PENILAIAN / ASESMEN
                - Asesmen Awal
                - Asesmen Formatif
                - Asesmen Sumatif
                
                ## LAMPIRAN
                - LAMPIRAN I: ASESMEN (Rubrik Sikap, Pengetahuan, dan Soal HOTS tentang {topik})
                - LAMPIRAN II: MATERI AJAR (Ringkasan materi {topik})
                - LAMPIRAN III: LKPD (Lembar Kerja Peserta Didik untuk topik {topik})
                """

                # Mengirim request ke NVIDIA API
                completion = client.chat.completions.create(
                    model="nvidia/nemotron-3-ultra-550b-a55b",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": f"Buat modul ajar lengkapnya sekarang."}
                    ],
                    temperature=0.7,
                    max_tokens=4000 
                )

                hasil_modul = completion.choices[0].message.content
                
                st.success("✅ Modul Ajar berhasil dibuat! Guru tinggal membacanya dan mengunduhnya.")
                
                # Tampilkan hasil
                st.markdown("---")
                st.markdown(hasil_modul)
                
                # Fitur Download untuk Guru
                st.download_button(
                    label="📥 Download Modul (Bisa dibuka di Word/Notepad)",
                    data=hasil_modul,
                    file_name=f"Modul_{mata_pelajaran}_{topik}.md",
                    mime="text/markdown"
                )
                
            except Exception as e:
                st.error(f"Terjadi kesalahan pada sistem AI: {e}")

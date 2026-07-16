# app.py
# AI Modul Ajar Generator v9.1 (Rian Dev) - NVIDIA NIM Edition
# Single File Deployment for Streamlit

import streamlit as st
import json
import re
import io
import os
from typing import Dict, List, Any, Optional, Literal
from dataclasses import dataclass, field, asdict
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ==========================================
# ⚙️ KONFIGURASI & SECRETS
# ==========================================
st.set_page_config(
    page_title="AI Modul Ajar Generator v9.1",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load Secrets (Streamlit Cloud / Local .streamlit/secrets.toml)
try:
    NVIDIA_API_KEY = st.secrets["NVIDIA_API_KEY"]
    NVIDIA_BASE_URL = st.secrets.get("NVIDIA_BASE_URL", "https://integrate.api.nvidia.com/v1")
    NVIDIA_MODEL = st.secrets.get("NVIDIA_MODEL", "meta/llama-3.1-70b-instruct")
except Exception:
    st.error("❌ **Secrets tidak ditemukan!** Masukkan `NVIDIA_API_KEY` di `.streamlit/secrets.toml` (local) atau Streamlit Cloud Settings.")
    st.stop()

# Init Client OpenAI-Compatible (NVIDIA NIM)
client = OpenAI(base_url=NVIDIA_BASE_URL, api_key=NVIDIA_API_KEY)

# ==========================================
# 🎨 STYLING & HELPERS DOCX
# ==========================================
COLOR_PRIMARY = RGBColor(0x1B, 0x3A, 0x5C)   # Dark Blue
COLOR_ACCENT = RGBColor(0x2E, 0x86, 0xC1)    # Blue
COLOR_LIGHT_BG = RGBColor(0xEB, 0xF5, 0xFB)  # Light Blue BG
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_BLACK = RGBColor(0x00, 0x00, 0x00)
COLOR_GREY = RGBColor(0x33, 0x33, 0x33)

def set_cell_shading(cell, color_hex: str):
    """Set background color cell tabel."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_heading_styled(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = COLOR_PRIMARY
        run.font.name = 'Calibri'
    return h

def add_normal(doc, text, bold=False, size=11, color=COLOR_BLACK, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = 'Calibri'
    return p

def add_bullet(doc, text, level=0, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(1.0 + level * 0.8)
    p.paragraph_format.space_after = Pt(2)
    if bold_prefix:
        run_b = p.add_run(bold_prefix)
        run_b.bold = True
        run_b.font.name = 'Calibri'
        run_b.font.size = Pt(11)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    return p

def make_table_pretty(table, header_color="1B3A5C"):
    """Style tabel: Header biru gelap putih, border, zebra stripe."""
    # Header
    for cell in table.rows[0].cells:
        set_cell_shading(cell, header_color)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.color.rgb = COLOR_WHITE
                run.font.bold = True
                run.font.size = Pt(10)
                run.font.name = 'Calibri'
    # Body
    for i, row in enumerate(table.rows[1:], start=1):
        for cell in row.cells:
            if i % 2 == 0:
                set_cell_shading(cell, "EBF5FB")
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(10)
                    run.font.name = 'Calibri'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

# ==========================================
# 📋 SCHEMA DATA (Pydantic-like Dict Structure)
# ==========================================
DEFAULT_TP = [
    "TP 1: Peserta didik mampu membaca kata dengan pola huruf beragam dan menjawab pertanyaan inferensial dari teks visual/audiovisual dengan tepat.",
    "TP 2: Peserta didik mampu menyajikan ide dengan pilihan kata yang bervariasi dan menulis teks eksplanasi secara sistematis berdasarkan fenomena atau tokoh inspiratif.",
    "TP 3: Peserta didik mampu mengevaluasi dan merefleksikan pengaruh tokoh-tokoh anak inspiratif terhadap kehidupan bermasyarakat.",
    "TP 4: Peserta didik mampu mengaitkan tindakan tokoh inspiratif dengan nilai-nilai Panca Cinta dan menerapkannya dalam kehidupan sehari-hari."
]

DEFAULT_PANCA_CINTA = [
    "Cinta Allah (Keimanan & Ketakwaan)",
    "Cinta Diri & Sesama (Adab & Empati)",
    "Cinta Ilmu (Bernalar Kritis & Kreatif)",
    "Cinta Lingkungan (Adab pada Alam)",
    "Cinta Tanah Air (Kebangsaan)"
]

DEFAULT_LINTAS_DISIPLIN = ["IPS", "IPA", "Pendidikan Pancasila", "Seni Budaya", "PAI"]

@dataclass
class MeetingPlan:
    nomor: int
    materi: str
    tp_fokus: str
    # Generated by AI
    opening: Dict = field(default_factory=dict)
    inti_pbl: List[Dict] = field(default_factory=list)
    closing: Dict = field(default_factory=dict)
    lkpd: Dict = field(default_factory=dict)
    asesmen_formatif: Dict = field(default_factory=dict)

@dataclass
class ModulAjarData:
    # Identitas
    mata_pelajaran: str = "Bahasa Indonesia"
    kelas: str = "VI"
    fase: str = "Fase C (Kelas 5-6)"
    semester: int = 1
    alokasi_jp: int = 24
    jumlah_pertemuan: int = 6
    durasi_menit: int = 35
    bab: str = "1"
    topik: str = "Anak-Anak yang Mengubah Dunia"
    model_pembelajaran: str = "PBL (Problem Based Learning)"
    metode: List[str] = field(default_factory=lambda: ["Ceramah Interaktif", "Diskusi Kelompok", "Tanya Jawab"])
    penyusun: str = "Mamik Muhapatin, S.Pd"
    sekolah: str = "MI. Miftahussalam"
    tahun_pelajaran: str = "2026/2027"
    # Kurikulum
    cp_kode: str = "CP 2025 No.46"
    cp_deskripsi: str = "Pada akhir Fase C, peserta didik memiliki kemampuan berbahasa untuk berkomunikasi dan bernalar sesuai tujuan, konteks sosial, dan akademis. Peserta didik mampu memahami, mengolah, dan menginterpretasikan informasi paparan tentang topik yang beragam dan karya sastra. Peserta didik mampu berpartisipasi aktif dalam diskusi, mempresentasikan, dan menanggapi informasi nonfiksi dan fiksi yang dipaparkan, serta menulis berbagai teks untuk menyampaikan pengamatan dan pengalamannya dengan lebih terstruktur dan jelas."
    tp_list: List[str] = field(default_factory=lambda: DEFAULT_TP)
    panca_cinta_fokus: List[str] = field(default_factory=lambda: DEFAULT_PANCA_CINTA)
    lintas_disiplin: List[str] = field(default_factory=lambda: DEFAULT_LINTAS_DISIPLIN)
    # Identifikasi PD
    pengetahuan_awal: str = "• Membaca teks sederhana dengan kelancaran dan pemahaman dasar\n• Menulis kalimat sederhana dengan ejaan yang benar\n• Mengenal jenis-jenis teks (narasi, deskripsi) secara dasar\n• Memahami kata-kata umum dalam kosakata Bahasa Indonesia"
    minat_belajar: str = "• Kisah inspiratif dan tokoh-tokoh berpengaruh di dunia\n• Kegiatan kolaboratif dan diskusi kelompok\n• Media visual dan audiovisual yang menarik\n• Proyek kreatif yang memberikan kesempatan berekspresi"
    latar_belakang: str = "Peserta didik kelas VI hidup di era digital dengan akses luas terhadap informasi. Mereka sering menjumpai kisah-kisah inspiratif dari berbagai tokoh anak muda di media sosial, televisi, maupun lingkungan sekitar."
    kebutuhan_belajar: str = "• Visualisasi: Gambar, foto, video pendek tentang tokoh anak inspiratif\n• Kinestetik: Aktivitas menulis, diskusi, bermain peran sebagai tokoh\n• Audio: Rekaman cerita, podcast inspiratif, atau pembacaan teks nyaring\n• Pendampingan khusus bagi peserta didik dengan kesulitan membaca melalui teks besar, audio support, dan bimbingan intensif"
    dpl: List[str] = field(default_factory=lambda: ["Penalaran Kritis", "Kreativitas", "Komunikasi", "Kolaborasi"])
    # Rencana Pertemuan (User Input)
    pertemuan_materi: List[str] = field(default_factory=lambda: [
        "Membaca Kata dengan Pola Huruf Beragam",
        "Menjawab Pertanyaan Inferensial dari Teks Visual/Audiovisual",
        "Menyajikan Ide dengan Pilihan Kata Bervariasi",
        "Menulis Teks Eksplanasi (Drafting & Peer Review)",
        "Menulis Teks Eksplanasi (Finalisasi & Presentasi)",
        "Refleksi, Evaluasi Akhir, & Penyelesaian Produk Akhir"
    ])
    pertemuan_tp_fokus: List[str] = field(default_factory=lambda: [
        "TP 1 (Memahami)", "TP 1 (Mengaplikasikan)", "TP 2 (Memahami)", "TP 2 & 3 (Mengaplikasikan)", "TP 3 & 4 (Merefleksikan)", "TP 1-4 (Evaluasi Sumatif)"
    ])
    # Generated Global
    asesmen_awal: Dict = field(default_factory=dict)
    asesmen_sumatif: Dict = field(default_factory=dict)
    remedial: Dict = field(default_factory=dict)
    pengayaan: Dict = field(default_factory=dict)
    glosarium: List[Dict] = field(default_factory=list)
    refleksi_guru: Dict = field(default_factory=dict)
    # Meetings Detail
    meetings: List[MeetingPlan] = field(default_factory=list)

# ==========================================
# 🤖 AI ENGINE (NVIDIA NIM)
# ==========================================
SYSTEM_PROMPT = """
Anda adalah **Lagos AI 9.1 (Rian Dev)**, Arsitek Kurikulum Merdeka & Deep Learning Specialist.
Tugas: Menghasilkan konten **Modul Ajar** format JSON yang **siap cetak**, pedagogis kuat, dan **kaya konteks lokal Indonesia**.

PRINSIP WAJIB (Deep Learning - Kemendikdasmen):
1. MINDFUL (Berkesadaran): Refleksi diam, observasi sadar, kesadaran proses belajar, "Menghadirkan Diri Sepenuhnya".
2. MEANINGFUL (Bermakna): Kontekstual kehidupan nyata, Panca Cinta, Profil Pelajar Pancasila, Life-relevance.
3. JOYFUL (Menggembirakan): Gamifikasi, kolaborasi menyenangkan, apresiasi, suasana positif, "Belajar itu Menyenangkan".

FORMAT OUTPUT: **HANYA JSON VALID**. TANPA MARKDOWN ```json ```. LANGSUNG OBJEK.
BAHASA: Indonesia Baku (EYD V).
GAYA: Profesional Guru, Detail, Spesifik (bukan generik), Rujuk Buku Siswa Kelas VI Bab 1.
"""

def call_llm(prompt: str, temperature=0.4, max_tokens=4096) -> Dict:
    """Panggil NVIDIA NIM via OpenAI Client, return parsed JSON."""
    try:
        response = client.chat.completions.create(
            model=NVIDIA_MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            temperature=temperature,
            max_tokens=max_tokens,
            response_format={"type": "json_object"} # NVIDIA NIM supports this mostly
        )
        content = response.choices[0].message.content
        # Cleaning just in case
        content = re.sub(r'^```json\s*|\s*```$', '', content.strip(), flags=re.MULTILINE)
        return json.loads(content)
    except json.JSONDecodeError as e:
        st.error(f"❌ AI Response bukan JSON valid: {e}\n\nRaw: {content[:500]}")
        return {}
    except Exception as e:
        st.error(f"❌ Error API NVIDIA: {e}")
        return {}

def generate_meeting_content(data: ModulAjarData, meeting: MeetingPlan) -> Dict:
    prompt = f"""
    KONTEKS MODUL AJAR:
    - Mapel: {data.mata_pelajaran} Kelas {data.kelas} {data.fase}
    - Semester {data.semester}, Bab {data.bab}: {data.topik}
    - Model: {data.model_pembelajaran} | Metode: {', '.join(data.metode)}
    - CP: {data.cp_kode}
    - Panca Cinta Fok

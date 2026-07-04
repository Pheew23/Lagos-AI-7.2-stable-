import streamlit as st
import os
from openai import OpenAI
import io
import re
from docx import Document
from dotenv import load_dotenv

# --- 1. KONFIGURASI KEAMANAN & ENVIRONMENT ---
load_dotenv() # Muat file .env
BASE_URL = "https://integrate.api.nvidia.com/v1"
# Aman dari hardcoding, ambil dari environment variable
NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY") 
MODEL_NAME = "z-ai/glm-5.2"

if not NVIDIA_API_KEY:
    st.error("API Key tidak ditemukan! Masukkan ke file .env")
    st.stop()

client = OpenAI(base_url=BASE_URL, api_key=NVIDIA_API_KEY)

# --- 2. FUNGSI EKSPOR WORD (DIPERBAIKI) ---
def buat_file_word(riwayat_pesan):
    doc = Document()
    doc.add_heading('Draf Hasil Kerja AI - GLM-5.2 Workspace', level=0)

    for msg in riwayat_pesan:
        # Gunakan flag is_hidden alih-alih string matching
        if msg.get("role") == "system" or msg.get("is_hidden"):
            continue

        if msg["role"] == "user":
            doc.add_heading("Pertanyaan / Instruksi Anda:", level=2)
            doc.add_paragraph(msg["content"])
        elif msg["role"] == "assistant":
            doc.add_heading("Jawaban AI:", level=2)
            paragraf_list = msg["content"].split('\n')
            for p_text in paragraf_list:
                if not p_text.strip(): continue
                match_heading = re.match(r'^(#{1,6})\s+(.*)$', p_text.strip())
                if match_heading:
                    level_pagar = len(match_heading.group(1))
                    teks_judul_bersih = match_heading.group(2).replace('**', '')
                    doc.add_heading(teks_judul_bersih, level=min(level_pagar, 3))
                    continue
                p = doc.add_paragraph()
                parts = re.split(r'(\*\*.*?\*\*)', p_text)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        p.add_run(part.replace('**', '')).bold = True
                    else:
                        p.add_run(part)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 3. INIT STATE DENGAN SLIDING WINDOW ---
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "system", "content": "Anda adalah GLM-5.2..."}
    ]

# --- 4. UI & LOGIC ---
st.set_page_config(page_title="GLM-5.2 Workspace", page_icon="🔮", layout="wide")

with st.sidebar:
    st.title("🔮 Kontrol AI")
    if st.button("🗑️ Reset Memori"):
        del st.session_state["messages"]
        st.rerun()
    if len(st.session_state.messages) > 1:
        st.download_button("📥 Download Word", data=buat_file_word(st.session_state.messages), file_name="Draf_AI.docx")

st.title("🔮 Lagos AI 7.7 (GLM-5.2)")

# Tampilkan riwayat
for msg in st.session_state.messages:
    if msg["role"] != "system" and not msg.get("is_hidden"):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

if user_input := st.chat_input("Ketik perintah..."):
    st.session_state.messages.append({"role": "user", "content": user_input})

    # Optimasi: Hanya kirim 10 pesan terakhir ke API untuk menghemat token
    context_window = st.session_state.messages[-10:]

    with st.chat_message("assistant"):
        try:
            response_stream = client.chat.completions.create(
                model=MODEL_NAME, messages=context_window, stream=True
            )
            def teks_generator():
                for chunk in response_stream:
                    if chunk.choices and (content := chunk.choices[0].delta.content):
                        yield content

            full_response = st.write_stream(teks_generator())
            st.session_state.messages.append({"role": "assistant

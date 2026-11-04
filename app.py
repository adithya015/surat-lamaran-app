import base64
from io import BytesIO
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import openpyxl
import pandas as pd
import requests
import streamlit as st
from xhtml2pdf import pisa

# Safe Import PDF Merger
try:
  from pypdf import PdfMerger
except ImportError:
  try:
    from PyPDF2 import PdfMerger
  except ImportError:
    PdfMerger = None

st.set_page_config(
    page_title="Generator Surat & CV", page_icon="📝", layout="centered"
)

# Link URL Google Apps Script Web App Anda
GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwldQWOaFEj_eASZ7tRrYkBfA5hZkAblfB6S4Z7PKFFqCukEpOD0ZBTOwMNLFUazruJBw/exec"


# Fungsi Ter-cache agar aplikasi tidak freeze/loading muter lama
@st.cache_data(ttl=15, show_spinner=False)
def ambil_data_sheets(url):
  try:
    response = requests.get(url, allow_redirects=True, timeout=5)
    if response.status_code == 200:
      return response.json()
  except Exception:
    return None
  return None


# Session State untuk Preview & Rekap Data
if "pdf_ready" not in st.session_state:
  st.session_state.pdf_ready = False
if "final_pdf_data" not in st.session_state:
  st.session_state.final_pdf_data = None
if "nama_pdf" not in st.session_state:
  st.session_state.nama_pdf = ""
if "email_tujuan" not in st.session_state:
  st.session_state.email_tujuan = ""
if "subjek_email" not in st.session_state:
  st.session_state.subjek_email = ""
if "body_email" not in st.session_state:
  st.session_state.body_email = ""

tab1, tab2 = st.tabs(
    ["📝 Buat Surat & Preview Email", "📊 Rekap Data Lamaran"]
)

# Data Profil Pelamar (Sesuai CV)
nama_pelamar = "Adithya Marhaendra Kusuma"
alamat_pelamar = "Perumahan Mutiara Citra Asri Blok D4 No 6 Candi Sidoarjo"
telepon = "082131009200"
email_pelamar = "adit.marhaendra@gmail.com"
pendidikan = "S2 Teknologi Informasi - ISTTS"
GMAIL_APP_PASSWORD_DEFAULT = "ssav gwlb tjwz exph"


def dapatkan_html_ttd():
  opsi = [
      "ttd_hd_white.png",
      "ttd_hd_transparent.png",
      "ttd.png",
      "ttd.jpg",
      "ttd.PNG",
  ]
  for f in opsi:
    if os.path.exists(f):
      with open(f, "rb") as img_file:
        b64 = base64.b64encode(img_file.read()).decode()
      return f'<img src="data:image/png;base64,{b64}" style="max-height: 55px; margin: 2px 0;">'
  return '<div style="height:55px;">[Tanda Tangan]</div>'


# ================= TAB 1: FORM SURAT & PREVIEW EMAIL =================
with tab1:
  st.subheader("1. Detail Lamaran Kerja")

  uploaded_cv = st.file_uploader(
      "📄 Upload File CV PDF Baru (Opsional)", type=["pdf"]
  )

  with st.form("form_surat"):
    perusahaan = st.text_input(
        "Nama Perusahaan", value="PT Mayora Indah Tbk (Plant Purwosari)"
    )
    lokasi = st.text_input("Lokasi / Kota", value="Purwosari, Pasuruan")
    posisi = st.text_input("Posisi Dilamar", value="Warehouse Unit Head")
    tanggal = st.text_input("Tanggal Surat", value="29 Agustus 2026")

    st.divider()
    st.subheader("2. Tujuan Pengiriman Email")

    email_tujuan = st.text_input(
        "Email Tujuan (HRD)", value="RECRUITMENT.PWS@MAYORA.CO.ID"
    )
    domisili_subjek = st.text_input(
        "Domisili Anda (Untuk Subjek Email)", value="SIDOARJO"
    )

    gmail_app_password = st.text_input(
        "🔑 Gmail App Password",
        value=GMAIL_APP_PASSWORD_DEFAULT,
        type="password",
    )

    submit_generate = st.form_submit_button("🔍 Generate PDF & Preview Email")

  if submit_generate:
    if not perusahaan or not posisi:
      st.error("Mohon isi Nama Perusahaan dan Posisi!")
    else:
      posisi_code = (
          "UH WAREHOUSE" if "WAREHOUSE" in posisi.upper() else posisi.upper()
      )
      subjek_draft = f"{posisi_code}_{domisili_subjek.upper()}"

      body_draft = f"""Yth. Tim Recruitment / HRD {perusahaan}
di {lokasi}

Dengan hormat,

Sehubungan dengan informasi lowongan pekerjaan yang dibuka oleh {perusahaan}, melalui email ini saya bermaksud mengajukan diri untuk mengisi posisi {posisi}.

Saya merupakan lulusan {pendidikan} dan S1 Teknik Informatika dengan latar belakang utama di bidang IT & Application Support selama lebih dari 8 tahun. Berkarir di PT Indosari Niaga Nusantara (Sari Roti Group), fokus utama saya adalah mengelola aplikasi operasional bisnis, melakukan validasi sistem, serta memastikan keandalan data untuk mendukung kelancaran proses bisnis.

Keahlian utama saya berfokus pada integrasi teknologi IT dengan kebutuhan operasional lintas divisi:
- IT & Application Support: Troubleshooting sistem/aplikasi operasional, konfigurasi master data (data hygiene), pembuatan SOP, serta edukasi/pelatihan teknis pengguna.
- Data Control & Analytics: Pengolahan data analitis (Excel Advanced, SQL, Python), sinkronisasi data antar-sistem (Distributor, Agen, Gudang), serta pengawasan fulfillment dan retur.
- Bridging IT & Business Process: Pengalaman menjembatani sistem IT dengan kebutuhan Supply Chain, pergudangan (inventory & stock opname), internal audit data, hingga pencatatan transaksi penjualan/sales.

Atas kontribusi dalam menyelaraskan sistem teknologi dengan efisiensi operasional, saya dianugerahi penghargaan 'Best SCM Indonesia 2025 - Plant Pasuruan'. Dengan fondasi IT yang kuat serta fleksibilitas untuk berkembang di area pergudangan, SCM, audit internal, maupun business development, saya meyakini dapat memberikan dampak positif bagi {perusahaan}.

Sebagai bahan pertimbangan, bersama email ini saya lampirkan berkas Surat Lamaran Kerja dan Curriculum Vitae (CV) dalam format PDF terpadu. Saya bersedia mengikuti tahapan interview secara onsite/offline maupun sistem kerja shift.

Besar harapan saya untuk diberikan kesempatan wawancara agar dapat mendiskusikan kualifikasi saya secara lebih rinci. Atas perhatian Bapak/Ibu, saya ucapkan terima kasih.

Hormat saya,
{nama_pelamar}
📞 {telepon} | ✉️ {email_pelamar}
📍 {alamat_pelamar}"""

      # A. Generate PDF Surat
      html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    @page {{ size: A4; margin: 15mm 20mm 15mm 20mm; }}

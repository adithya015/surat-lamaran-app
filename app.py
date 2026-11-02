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

# Session State untuk Preview & Simpan Status
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

# URL Google Apps Script Web App
GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwldQWOaFEj_eASZ7tRrYkBfA5hZkAblfB6S4Z7PKFFqCukEpOD0ZBTOwMNLFUazruJBw/exec"


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

    submit_generate = st.form_submit_button(
        "🔍 Generate PDF & Preview Email"
    )

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
- IT & Application Support: Troubleshooting sistem/aplikasi operasional,

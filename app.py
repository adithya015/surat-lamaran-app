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

# Session State untuk Preview & Rekap
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

# Link URL Google Apps Script Web App
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
                    body {{ font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.35; color: #000; }}
                    p {{ margin: 6px 0; }}
                    .header {{ text-align: right; margin-bottom: 12px; }}
                    .recipient {{ margin-bottom: 15px; }}
                    .content {{ text-align: justify; }}
                    .bio-table {{ width: 100%; margin-left: 15px; margin-top: 4px; margin-bottom: 8px; }}
                    .bio-table td {{ padding: 1px 0; vertical-align: top; }}
                    .bio-table td.label {{ width: 26%; }}
                    .bio-table td.colon {{ width: 3%; }}
                    ol {{ margin-top: 2px; margin-bottom: 6px; padding-left: 22px; }}
                    ol li {{ padding: 1px 0; }}
                    .signature-table {{ width: 100%; margin-top: 15px; }}
                </style>
            </head>
            <body>
                <div class="header">Surabaya, {tanggal}</div>
                <div class="recipient">Yth. Bapak/Ibu HRD<br><strong>{perusahaan}</strong><br>di {lokasi}</div>
                <div class="content">
                    <p>Dengan hormat,</p>
                    <p>Berdasarkan informasi lowongan pekerjaan yang saya peroleh, bahwa <strong>{perusahaan}</strong> sedang membuka lowongan pekerjaan. Melalui surat ini saya bermaksud menyampaikan ketertarikan saya untuk melamar pekerjaan pada posisi <strong>{posisi}</strong>.</p>
                    <p>Berikut adalah biodata singkat saya:</p>
                    <table class="bio-table">
                        <tr><td class="label">Nama</td><td class="colon">:</td><td>{nama_pelamar}</td></tr>
                        <tr><td class="label">Pendidikan Terakhir</td><td class="colon">:</td><td>{pendidikan}</td></tr>
                        <tr><td class="label">Alamat Domisili</td><td class="colon">:</td><td>{alamat_pelamar}</td></tr>
                        <tr><td class="label">No. Telepon / WA</td><td class="colon">:</td><td>{telepon}</td></tr>
                        <tr><td class="label">Email</td><td class="colon">:</td><td>{email_pelamar}</td></tr>
                    </table>
                    <p>Saat ini saya dalam kondisi kesehatan yang sangat baik dan siap untuk bekerja keras serta berkontribusi positif bagi perusahaan. Saya memiliki latar belakang pendidikan S2 Teknologi Informasi serta pengalaman kerja lebih dari 8 tahun di bidang IT, SCM, dan Data Control yang saya yakini dapat mendukung kinerja operasional pada posisi tersebut.</p>
                    <p>Sebagai bahan pertimbangan Bapak/Ibu, bersama surat ini turut saya lampirkan:</p>
                    <ol>
                        <li>Curriculum Vitae (CV)</li>
                        <li>Fotokopi Ijazah Terakhir dan Transkrip Nilai</li>
                        <li>Portofolio / Dokumen Pendukung</li>
                        <li>Pasfoto Terbaru</li>
                    </ol>
                    <p>Besar harapan saya agar Bapak/Ibu bersedia meluangkan waktu untuk memberikan kesempatan wawancara. Atas perhatian dan kesempatan yang diberikan, saya ucapkan terima kasih.</p>
                </div>
                <table class="signature-table">
                    <tr>
                        <td style="width: 60%;"></td>
                        <td style="width: 40%;">
                            <p style="margin:0;">Hormat saya,</p>
                            {dapatkan_html_ttd()}
                            <p style="text-decoration: underline; font-weight: bold; margin-top: 0;">{nama_pelamar}</p>
                        </td>
                    </tr>
                </table>
            </body>
            </html>
            """
      surat_buffer = BytesIO()
      pisa.CreatePDF(html_content, dest=surat_buffer)
      surat_buffer.seek(0)

      # B. Merge Surat + CV
      if PdfMerger is not None:
        merger = PdfMerger()
        merger.append(surat_buffer)

        if uploaded_cv is not None:
          merger.append(BytesIO(uploaded_cv.read()))
        elif os.path.exists("cv.pdf"):
          merger.append("cv.pdf")
        elif os.path.exists("CV Adithya Marhaendra Kusuma ats.pdf"):
          merger.append("CV Adithya Marhaendra Kusuma ats.pdf")

        final_buffer = BytesIO()
        merger.write(final_buffer)
        merger.close()
        final_pdf_data = final_buffer.getvalue()
      else:
        final_pdf_data = surat_buffer.getvalue()

      nama_pdf = f"Surat_Lamaran_dan_CV_{perusahaan.replace(' ', '_')}.pdf"

      # C. Simpan ke Session State
      st.session_state.pdf_ready = True
      st.session_state.final_pdf_data = final_pdf_data
      st.session_state.nama_pdf = nama_pdf
      st.session_state.email_tujuan = email_tujuan
      st.session_state.subjek_email = subjek_draft
      st.session_state.body_email = body_draft
      st.session_state.app_password = gmail_app_password
      st.session_state.tanggal = tanggal
      st.session_state.perusahaan = perusahaan
      st.session_state.lokasi = lokasi
      st.session_state.posisi = posisi

  # ================= AREA PREVIEW & PENGIRIMAN =================
  if st.session_state.pdf_ready:
    st.divider()
    st.success(
        "✅ Dokumen PDF Berhasil Dibuat! Periksa atau edit draft email di bawah"
        " sebelum dikirim."
    )

    col_dl, col_blank = st.columns([1, 1])
    with col_dl:
      st.download_button(
          label="📥 Download & Cek File PDF",
          data=st.session_state.final_pdf_data,
          file_name=st.session_state.nama_pdf,
          mime="application/pdf",
          use_container_width=True,
      )

    st.subheader("📧 Form Preview & Edit Email")
    final_to = st.text_input(
        "Email Tujuan (HRD):", value=st.session_state.email_tujuan
    )
    final_subject = st.text_input(
        "Subjek Email:", value=st.session_state.subjek_email
    )
    final_body = st.text_area(
        "Isi Body Email (Bisa Diedit):",
        value=st.session_state.body_email,
        height=300,
    )

    if st.button("📤 Kirim Email Sekarang", type="primary"):
      if not st.session_state.app_password:
        st.error("Gmail App Password belum terisi!")
      else:
        try:
          msg = MIMEMultipart()
          msg["From"] = email_pelamar
          msg["To"] = final_to
          msg["Subject"] = final_subject
          msg.attach(MIMEText(final_body, "plain"))

          part = MIMEBase("application", "octet-stream")
          part.set_payload(st.session_state.final_pdf_data)
          encoders.encode_base64(part)
          part.add_header(
              "Content-Disposition",
              f"attachment; filename={st.session_state.nama_pdf}",
          )
          msg.attach(part)

          server = smtplib.SMTP("smtp.gmail.com", 587)
          server.starttls()
          server.login(
              email_pelamar, st.session_state.app_password.replace(" ", "")
          )
          server.sendmail(email_pelamar, final_to, msg.as_string())
          server.quit()

          # Rekap HANYA saat email berhasil dikirim
          try:
            payload = {
                "tanggal": st.session_state.get("tanggal", ""),
                "perusahaan": st.session_state.get("perusahaan", ""),
                "lokasi": st.session_state.get("lokasi", ""),
                "posisi": st.session_state.get("posisi", ""),
                "status": "Terkirim",
                "nama_pdf": st.session_state.nama_pdf,
            }
            requests.post(GOOGLE_SHEET_URL, json=payload, timeout=5)
          except Exception:
            pass

          st.balloons()
          st.success(
              f"🚀 Email & Berkas Lampiran PDF Berhasil Terkirim ke {final_to}!"
          )
        except Exception as e:
          st.error(f"Gagal mengirim email: {e}")

# ================= TAB 2: REKAP DATA FROM GOOGLE SHEETS =================
with tab2:
  st.subheader("📊 Histori Lamaran Kerja (Google Sheets)")
  try:
    response = requests.get(GOOGLE_SHEET_URL, allow_redirects=True, timeout=10)
    if response.status_code == 200:
      try:
        res = response.json()
        if isinstance(res, list) and len(res) > 1:
          df = pd.DataFrame(res[1:], columns=res[0])

          # 1. Filter out 'Draft Prepared' agar tidak ada data ganda
          if "Status" in df.columns:
            df = df[df["Status"] != "Draft Prepared"]

          # 2. Format Tanggal agar ringkas (hanya YYYY-MM-DD)
          if "Tanggal" in df.columns:
            df["Tanggal"] = df["Tanggal"].astype(str).apply(
                lambda x: x.split("T")[0] if "T" in x else x
            )

          # 3. Reset index agar urutan nomor kembali 1, 2, 3, 4, 5...
          df = df.reset_index(drop=True)
          df.index = df.index + 1

          st.metric(label="Total Lamaran Terkirim", value=len(df))
          st.dataframe(df, use_container_width=True)
        else:
          st.info(
              "Belum ada data lamaran yang dibuat di Google Sheets (atau sheet"
              " masih kosong)."
          )
      except Exception:
        st.error(
            "⚠️ Akses Google Apps Script ditolak/tidak valid. Pastikan"
            " aksesibilitas Web App pada Google Apps Script sudah diset ke"
            " 'Anyone' (Siapa saja)."
        )
    else:
      st.error(
          f"Gagal terhubung ke Google Sheets (Status HTTP:"
          f" {response.status_code})."
      )
  except Exception as e:
    st.error(f"Gagal mengambil data rekap dari Google Sheets: {e}")

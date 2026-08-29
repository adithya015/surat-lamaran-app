import base64
from io import BytesIO
import os
import openpyxl
import pandas as pd
import streamlit as st
from xhtml2pdf import pisa

# Safe Import untuk PDF Merger (Anti-Crash jika library baru dipasang)
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

# Tab Menu Navigasi
tab1, tab2 = st.tabs(["📝 Buat Surat + Merge CV", "📊 Rekap Data Lamaran"])

# Data Profil Pelamar (Default)
nama_pelamar = "Adithya Marhaendra Kusuma"
alamat_pelamar = "Perumahan Mutiara Citra Asri Blok D4 No 6 Candi Sidoarjo"
telepon = "082131009200"
email = "adit.marhaendra@gmail.com"
pendidikan = "S2 Teknologi Informasi - ISTTS"
file_excel = "data_lamaran.xlsx"


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


# ================= TAB 1: FORM SURAT & MERGE CV =================
with tab1:
  st.subheader("Form Lamaran Kerja")

  uploaded_cv = st.file_uploader(
      "📄 Upload File CV PDF Baru (Opsional)", type=["pdf"]
  )

  with st.form("form_surat"):
    perusahaan = st.text_input(
        "Nama Perusahaan", placeholder="PT Example Indonesia"
    )
    lokasi = st.text_input("Lokasi / Kota", placeholder="Surabaya")
    posisi = st.text_input("Posisi Dilamar", placeholder="IT Manager")
    tanggal = st.text_input("Tanggal Surat", placeholder="29 Agustus 2026")
    submit = st.form_submit_button("🚀 Generate PDF Lengkap (Surat + CV)")

  if submit:
    if not perusahaan or not posisi:
      st.error("Mohon isi Nama Perusahaan dan Posisi!")
    else:
      # 1. Generate Surat Lamaran (HTML -> PDF 1 Halaman)
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
                        <tr><td class="label">Email</td><td class="colon">:</td><td>{email}</td></tr>
                    </table>
                    <p>Saat ini saya dalam kondisi kesehatan yang sangat baik dan siap untuk bekerja keras serta berkontribusi positif bagi perusahaan. Saya memiliki latar belakang pendidikan dan pengalaman kerja selama lebih dari 8 tahun di bidang IT, SCM, dan Data Control yang saya yakini dapat mendukung kinerja saya pada posisi tersebut.</p>
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

      # 2. Penggabungan PDF (Surat + CV)
      cv_terpakai = False

      if PdfMerger is None:
        st.warning(
            "⚠️ Library PDF Merger belum aktif di server. Menampilkan Surat"
            " Lamaran saja."
        )
        final_pdf_data = surat_buffer.getvalue()
      else:
        merger = PdfMerger()
        merger.append(surat_buffer)

        # Urutan prioritas pencarian file CV
        if uploaded_cv is not None:
          merger.append(BytesIO(uploaded_cv.read()))
          cv_terpakai = True
        elif os.path.exists("cv.pdf"):
          merger.append("cv.pdf")
          cv_terpakai = True
        elif os.path.exists("CV Adithya Marhaendra Kusuma ats.pdf"):
          merger.append("CV Adithya Marhaendra Kusuma ats.pdf")
          cv_terpakai = True

        final_buffer = BytesIO()
        merger.write(final_buffer)
        merger.close()
        final_pdf_data = final_buffer.getvalue()

      # 3. Simpan Log Data ke Excel
      if os.path.exists(file_excel):
        wb = openpyxl.load_workbook(file_excel)
        ws = (
            wb["Daftar Lamaran"]
            if "Daftar Lamaran" in wb.sheetnames
            else wb.active
        )
      else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Daftar Lamaran"
        ws.append([
            "Tanggal",
            "Nama Perusahaan",
            "Lokasi",
            "Posisi",
            "Status",
            "Nama File PDF",
        ])

      nama_pdf = f"Surat_Lamaran_dan_CV_{perusahaan.replace(' ', '_')}.pdf"
      ws.append([tanggal, perusahaan, lokasi, posisi, "Terkirim", nama_pdf])
      wb.save(file_excel)

      # Notifikasi & Tombol Download
      if cv_terpakai:
        st.success("✅ Surat Lamaran + CV Berhasil Digabungkan!")
      else:
        st.info(
            "ℹ️ Surat Lamaran Berhasil Dibuat (File CV belum terdeteksi di"
            " server / di-upload)."
        )

      st.download_button(
          label="📥 Download Paket Lamaran PDF",
          data=final_pdf_data,
          file_name=nama_pdf,
          mime="application/pdf",
      )

# ================= TAB 2: REKAP DATA =================
with tab2:
  st.subheader("📊 Histori Lamaran Kerja")
  if os.path.exists(file_excel):
    df = pd.read_excel(file_excel)
    st.metric(label="Total Lamaran Dibuat", value=len(df))
    st.dataframe(df, use_container_width=True)

    with open(file_excel, "rb") as f:
      st.download_button(
          label="📥 Download Rekap Excel (.xlsx)",
          data=f.read(),
          file_name="Rekap_Lamaran_Kerja.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )
  else:
    st.info("Belum ada data lamaran yang dibuat.")

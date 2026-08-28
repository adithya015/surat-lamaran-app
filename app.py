import base64
from io import BytesIO
import os
import openpyxl
import streamlit as st
from xhtml2pdf import pisa

st.set_page_config(
    page_title="Generator Surat Lamaran", page_icon="📝", layout="centered"
)
st.title("📝 Generator Surat Lamaran PDF")

# Form Input di Browser
with st.form("form_surat"):
  perusahaan = st.text_input(
      "Nama Perusahaan", placeholder="PT Example Indonesia"
  )
  lokasi = st.text_input("Lokasi / Kota", placeholder="Surabaya")
  posisi = st.text_input("Posisi Dilamar", placeholder="IT Manager")
  tanggal = st.text_input("Tanggal Surat", placeholder="28 Agustus 2026")

  submit = st.form_submit_button("🚀 Buat Surat PDF")

if submit:
  if not perusahaan or not posisi:
    st.error("Mohon isi Nama Perusahaan dan Posisi!")
  else:
    nama_pelamar = "Adithya Marhaendra Kusuma"
    alamat_pelamar = (
        "Perumahan Mutiara Citra Asri Blok D4 No 6 Candi Sidoarjo"
    )
    telepon = "082131009200"
    email = "adit.marhaendra@gmail.com"
    pendidikan = "S2 Teknologi Informasi - ISTTS"

    def dapatkan_html_ttd():
      opsi = ["ttd_hd_white.png", "ttd.png", "ttd.jpg"]
      for f in opsi:
        if os.path.exists(f):
          with open(f, "rb") as img_file:
            b64 = base64.b64encode(img_file.read()).decode()
          return f'<img src="data:image/png;base64,{b64}" style="max-height: 55px; margin: 2px 0;">'
      return '<div style="height:55px;">[Tanda Tangan]</div>'

    html_content = f"""
         <!DOCTYPE html>
         <html>
         <head>
             <style>
                 @page {{ size: A4; margin: 15mm 20mm 15mm 20mm; }}
                 body {{ font-family: 'Times New Roman', serif; font-size: 11pt; line-height: 1.35; }}
                 p {{ margin: 6px 0; }}
                 .header {{ text-align: right; margin-bottom: 12px; }}
                 .recipient {{ margin-bottom: 15px; }}
                 .content {{ text-align: justify; }}
                 .bio-table {{ width: 100%; margin-left: 15px; margin-top: 4px; margin-bottom: 8px; }}
                 .bio-table td {{ padding: 1px 0; vertical-align: top; }}
                 .bio-table td.label {{ width: 26%; }}
                 .bio-table td.colon {{ width: 3%; }}
                 ol {{ margin-top: 2px; margin-bottom: 6px; padding-left: 22px; }}
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

    # Render PDF ke Memory Buffer
    pdf_buffer = BytesIO()
    pisa.CreatePDF(html_content, dest=pdf_buffer)
    pdf_data = pdf_buffer.getvalue()

    # Record ke Excel
    file_excel = "data_lamaran.xlsx"
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

    nama_pdf = f"Surat_Lamaran_{perusahaan.replace(' ', '_')}.pdf"
    ws.append([tanggal, perusahaan, lokasi, posisi, "Selesai", nama_pdf])
    wb.save(file_excel)

    st.success("✅ Surat PDF Berhasil Dibuat!")

    # Tombol Download PDF Langsung
    st.download_button(
        label="📥 Download Surat PDF",
        data=pdf_data,
        file_name=nama_pdf,
        mime="application/pdf",
    )
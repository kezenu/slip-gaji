import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from Include.passwordsmtp import pw_smtp
from datetime import datetime
import os
import pandas as pd
import time

def bulan_indonesia(x):
    hari_ini = datetime.now()
    bulan_ind = [
        "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    if x == "bulan":
        return bulan_ind[hari_ini.month]
    elif x == "tahun":
        return f"{hari_ini.strftime('%Y')}"
    return None

class PENGIRIM_EMAIL:
    def __init__(self, nama_penerima, email_penerima, nama_file):
        self.nama_penerima = nama_penerima
        self.email_penerima = email_penerima
        self.nama_file = nama_file
    
    def template_body(self):
        return f"""
Kepada Yth. 
{self.nama_penerima}

Berikut terlampir slip gaji bulan {bulan_indonesia("bulan")} {bulan_indonesia("tahun")}.
Setiap file slip gaji di proteksi dengan NIK masing-masing.
Mohon diperiksa dan disimpan dengan baik.
Jika terdapat kesalahan atau keluhan, silakan hubungi bagian HRD.
Terima kasih.

PT. ALKATRA
Hormat kami,

Joel Simatupang
HRD
"""

    def template_email(self):
        email_pengirim = pw_smtp(2)
        app_password   = pw_smtp(1)
        email_penerima = self.email_penerima

        msg = MIMEMultipart()
        msg["From"] = email_pengirim
        msg["To"] = email_penerima
        msg["Subject"] = f"Slip Gaji - Periode Bulan {bulan_indonesia('bulan')} {bulan_indonesia('tahun')}"
        msg.attach(MIMEText(self.template_body(), "plain"))

        # Lampiran PDF
        with open(self.nama_file, "rb") as lampiran:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(lampiran.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename={os.path.basename(self.nama_file)}")
            msg.attach(part)

        try:
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            server.login(email_pengirim, app_password)
            server.sendmail(email_pengirim, email_penerima, msg.as_string())
            server.quit()
            print(f"✅ Slip gaji {self.nama_penerima} berhasil dikirim ke email {email_penerima}!")
        except Exception as e:
            print(f"❌ {self.nama_penerima} tidak terkirim, kode Error: ", e)


# --- kirim ---
def kirim_email(data):
    df = pd.DataFrame(data)
    top5 = df.iloc[:15]
    for row in top5.itertuples(index=False): #untuk percobaan 5 email teratas. kalau sudah oke ganti df
        nama = row.nama
        email = row.email
        if not email:
            print(f"❌ Slip gaji {nama} tidak terkirim, email tidak ditemukan")
            continue
        nama_file = os.path.join("data", "slip", f'{nama.replace(" ", "_")}.pdf')
        email_kirim = PENGIRIM_EMAIL(nama, email, nama_file)
        email_kirim.template_email()

        time.sleep(2)
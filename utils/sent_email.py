import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from Include.passwordsmtp import pw_smtp
from datetime import datetime
import os
from utils.slip_generator import df

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

def template_body():
    return f"""
Kepada Yth. 
Bapak/Ibu Karyawan PT. ALKATRA

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

class PENGIRIM_EMAIL:
    def __init__(self, penerima, nama_file):
        self.penerima = penerima
        self.nama_file = nama_file

    def template_email(self):
        email_pengirim = pw_smtp(2)
        app_password   = pw_smtp(1)
        email_penerima = self.penerima

        msg = MIMEMultipart()
        msg["From"] = email_pengirim
        msg["To"] = email_penerima
        msg["Subject"] = f"Slip Gaji - {bulan_indonesia('bulan')} {bulan_indonesia('tahun')}"
        msg.attach(MIMEText(template_body(), "plain"))

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
            print(f"✅ Email berhasil dikirim ke {email_penerima}!")
        except Exception as e:
            print("❌ Error:", e)


# --- Contoh kirim ---
def kirim_email():
    nama_file = os.path.join("data", "slip", "RAMLI.pdf")
    email_kirim = PENGIRIM_EMAIL("wsanjaya69@gmail.com", nama_file)
    email_kirim.template_email()

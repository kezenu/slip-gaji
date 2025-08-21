import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from Include.passwordsmtp import pw_smtp
from datetime import datetime

def bulan_indonesia(x):
    hari_ini = datetime.now()
    if x == "bulan":
        bulan_indonesia = [
            "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
        return bulan_indonesia[hari_ini.month]
    elif x == "tahun":
        return f"{hari_ini.strftime('%Y')}"
    else:
        return None
    
class PENGIRIM_EMAIL:
    def __init__(self, penerima):
        self.penerima = penerima

    def kirim_email(self):
        # Data login Gmail
        email_pengirim = pw_smtp(2)
        app_password =  pw_smtp(1) # App password dari langkah 2FA
        email_penerima = self.penerima

        # Buat pesan email
        msg = MIMEMultipart()
        msg["From"] = email_pengirim
        msg["To"] = email_penerima
        msg["Subject"] = f"Slip gaji"

        # Isi email
        body = f"""
    Kepada Yth. Bapak/Ibu Karyawan PT. ALKATRA

    Berikut terlampir slip gaji bulan {datetime.now()} [Tahun].
    Setiap file slip gaji di proteksi dengan NIK karyawan masing-masing.
    Mohon diperiksa dan disimpan dengan baik.
    Jika terdapat kesalahan atau keluhan, silakan hubungi bagian HRD.
    Terima kasih.

    Hormat saya,


    [Nama Pengirim/HRD]
    [Jabatan]
    [Perusahaan]

    """
        msg.attach(MIMEText(body, "plain"))

        try:
            # Koneksi ke SMTP Gmail (pakai SSL port 465)
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            server.login(email_pengirim, app_password)
            server.sendmail(email_pengirim, email_penerima, msg.as_string())
            server.quit()
            print("✅ Email berhasil dikirim!")
        except Exception as e:
            print("❌ Error:", e)

# email_kirim = PENGIRIM_EMAIL("wsanjaya69@gmail.com")
# email_kirim.kirim_email()

body = f"""
    Kepada Yth. 
    Bapak/Ibu Karyawan PT. ALKATRA

    Berikut terlampir slip gaji bulan {bulan_indonesia("bulan")} {bulan_indonesia("tahun")}.
    Setiap file slip gaji di proteksi dengan NIK masing-masing.
    Mohon diperiksa dan disimpan dengan baik.
    Jika terdapat kesalahan atau keluhan, silakan hubungi bagian HRD.
    Terima kasih.

    PT. ALKATRA
    Hormat saya,


    Joel Simatupang
    HRD

    """

print(body)
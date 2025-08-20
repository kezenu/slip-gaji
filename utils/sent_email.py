import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from Include.passwordsmtp import pw_smtp


# Data login Gmail
email_pengirim = pw_smtp(2)
app_password =  pw_smtp(1) # App password dari langkah 2
email_penerima = "wsanjaya69@gmail.com"

# Buat pesan email
msg = MIMEMultipart()
msg["From"] = email_pengirim
msg["To"] = email_penerima
msg["Subject"] = "Test SMTP Gmail"

# Isi email
body = "Halo, ini email percobaan via SMTP Gmail + Python!"
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

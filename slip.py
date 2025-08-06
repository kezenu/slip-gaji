import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm

# Baca data dan konversi ke list of dict
df = pd.read_excel('gaji_juli.xlsx')
df.fillna(0, inplace=True)
df.set_index(keys='nik', inplace=True)
data = df.to_dict(orient='records')

def multi_slip(data, filename="data/slip_gaji.pdf"):
    c = canvas.Canvas(filename, pagesize=[25 * cm, 15 * cm])
    width, height = 25 * cm, 15 * cm
    text_width = c.stringWidth("SLIP GAJI KARYAWAN", "Helvetica-Bold", 16)
    center = (24.13 * cm - text_width * cm ) / 2
    print(text_width)

    for item in data:
        c.setFont("Helvetica-Bold", 16)
        c.drawString(center * cm, height - 2 * cm, "SLIP GAJI KARYAWAN")

        c.setFont("Helvetica", 11)
        c.drawString(2 * cm, height - 3 * cm, f"NAMA{' '* 15}:")
        c.drawString(2 * cm, height - 3.5 * cm, f"LOKASI{' '* 15}:")
        c.drawString(7 * cm, height - 3 * cm, f"{item['nama']}")

        # ⛳️ Halaman baru untuk setiap slip
        c.showPage()  # ← Ini penting!

    c.save()

# Jalankan fungsi
multi_slip(data)

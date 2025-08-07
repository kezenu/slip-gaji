import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm

# Baca data dan konversi ke list of dict
df = pd.read_excel('gaji_contoh.xlsx')
df.fillna(0, inplace=True)
df.set_index(keys='nik', inplace=True)
data = df.to_dict(orient='records')
print(df.head(3))
def multi_slip(data, filename="data/slip_gaji.pdf"):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter

    for item in data:
        c.setFont("Helvetica-Bold", 16)
        c.drawString(8 * cm, height - 2 * cm, "SLIP GAJI KARYAWAN")

        #  Informasi pekerja
        c.setFont("Helvetica", 11)
        c.drawString(2 * cm, height - 3 * cm, f"Nama")
        c.drawString(2 * cm, height - 3.5 * cm, f"Lokasi")
        c.drawString(2 * cm, height - 4 * cm, f"Status")
        c.drawString(7 * cm, height - 3 * cm, f" : {item['nama']}")
        c.drawString(7 * cm, height - 3.5 * cm, f" : {item['lokasi']}")
        c.drawString(7 * cm, height - 4 * cm, f" : {item['status']}")
        

        c.drawString(2 * cm, height - 7 * cm, f"PENERIMAAN")
        c.drawString(2 * cm, height - 8 * cm, f"Gaji Pokok")
        c.drawString(7 * cm, height - 8 * cm, f" : Rp. {item['basic']:,.2f}")

        # ⛳️ Halaman baru untuk setiap slip
        c.showPage()  # ← Ini penting!

    c.save()

# Jalankan fungsi
multi_slip(data)

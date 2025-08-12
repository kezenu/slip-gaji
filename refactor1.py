import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from datetime import date

bulan_indonesia = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

# Baca data
df = pd.read_excel('gaji_contoh.xlsx')

# Mengisi semua data yang kosong dengan nilai 0
df.fillna(0, inplace=True)

# Menetapkan indek berdasarakn NIK
df.set_index(keys='nik', inplace=True)

# Convert semua data ke Integer jika ada
df[['basic', 'jabatan', 'tunjangan_bpjstk', 
    'tunjangan_bpjskes', 'cuti', 'holiday',
    'lembur', 'event', 'total_ot', 
    'revisi', 'thr', 'pph_kantor',
    'bruto', 'operasional', 'kelebihan', 
    'kekurangan_jam','potongan_bpjstk', 
    'pp21', 'total_potongan', 'netto']] = df[['basic', 'jabatan', 'tunjangan_bpjstk', 
                                              'tunjangan_bpjskes', 'cuti', 'holiday',
                                              'lembur', 'event', 'total_ot', 
                                              'revisi', 'thr', 'pph_kantor',
                                              'bruto', 'operasional', 'kelebihan', 
                                              'kekurangan_jam', 'potongan_bpjstk',
                                              'pp21', 'total_potongan','netto']].astype(int)

# Convert Dataframe ke Dict dan masukan nik
data = df.reset_index().to_dict(orient='records')
print(df.columns)

c = canvas.Canvas(filename="data/contoh.pdf", pagesize=A4)
width, height = A4


def render_header(c, item, x, start_y):
    text_width = c.stringWidth("SLIP GAJI KARYAWAN", "Helvetica-Bold", 16)
    center = (width - text_width) / 2
    c.setFont("Helvetica-Bold", 16)
    c.drawString(center, height - 1 * cm, "SLIP GAJI KARYAWAN")

    c.setFont("Helvetica", 11)
    c.drawString(x, start_y, ": Nama")
    c.drawString(x + 5 * cm, start_y, f"{item['nama']}")
    y = start_y - 0.5 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Lokasi")
    c.drawString(x + 5 * cm, y, f": {item['lokasi']}")
    y -= 0.5 * cm
    
    return y


def render_penerimaan(c, item, x, start_y):
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, start_y, "PENERIMAAN")
    y = start_y - 0.5 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Gaji Pokok")
    c.drawRightString(x + 12 * cm, y, f": Rp. {item['basic']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Tunjangan Jabatan")
    c.drawRightString(x + 12 * cm, y, f": Rp. {item['jabatan']:,.2f}")
    y -= 0.5 * cm

    # Tambah baris lain di sini
    c.drawString(x, y, "Lembur")
    c.drawRightString(x + 12 * cm, y, f": Rp. {item['lembur']:,.2f}")
    y -= 0.5 * cm

    return y  # supaya bagian selanjutnya bisa lanjut dari sini




def multi_slip(data, filename="data/slip_gaji1.pdf"):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    for item in data:
        margin_x = 2 * cm
        start_y = height - 3 * cm 
        render_header(c, item, margin_x, start_y)
        render_penerimaan(c, item, margin_x, render_header(c, item, margin_x, start_y))

        c.showPage()
    c.save()


multi_slip(data)

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

def render_penerimaan(c, item, x, y_start):
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y_start, "PENERIMAAN")
    y = y_start - 0.5 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Gaji Pokok")
    c.drawRightString(x + 12 * cm, y, f"Rp. {item['basic']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Tunjangan Jabatan")
    c.drawRightString(x + 12 * cm, y, f"Rp. {item['jabatan']:,.2f}")
    y -= 0.5 * cm

    # Tambah baris lain di sini tanpa takut
    c.drawString(x, y, "Lembur")
    c.drawRightString(x + 12 * cm, y, f"Rp. {item['lembur']:,.2f}")
    y -= 0.5 * cm

    return y  # supaya bagian selanjutnya bisa lanjut dari sini

def multi_slip(data, filename="data/slip_gaji1.pdf"):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    for item in data:
        render_penerimaan(c, item, width, height)
        c.showPage()  # slip baru per orang
    c.save()

multi_slip(data)
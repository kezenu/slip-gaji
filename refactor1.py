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
    'pph21', 'total_potongan', 'netto']] = df[['basic', 'jabatan', 'tunjangan_bpjstk', 
                                              'tunjangan_bpjskes', 'cuti', 'holiday',
                                              'lembur', 'event', 'total_ot', 
                                              'revisi', 'thr', 'pph_kantor',
                                              'bruto', 'operasional', 'kelebihan', 
                                              'kekurangan_jam', 'potongan_bpjstk',
                                              'pph21', 'total_potongan','netto']].astype(int)

# Convert Dataframe ke Dict dan masukan nik
data = df.reset_index().to_dict(orient='records')
print(df.columns)

c = canvas.Canvas(filename="data/contoh.pdf", pagesize=A4)
width, height = A4


def render_header(c, item, x, start_y):
    text_width = c.stringWidth("SLIP GAJI KARYAWAN", "Helvetica-Bold", 16)
    center = (width - text_width) / 2
    c.setFont("Helvetica-Bold", 16)
    c.drawString(center, height - 3 * cm, "SLIP GAJI KARYAWAN")

    c.setFont("Helvetica", 11)
    c.drawString(x, start_y, "Nama")
    c.drawString(x + 5 * cm, start_y, f": {item['nama']}")
    y = start_y - 0.5 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Lokasi")
    c.drawString(x + 5 * cm, y, f": {item['lokasi']}")
    y -= 0.5 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Status")
    c.drawString(x + 5 * cm, y, f": {item['status']}")
    y -= 0.5 * cm

    return y - 1 * cm


def render_penerimaan(c, item, x, start_y):
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, start_y, "PENERIMAAN")
    y = start_y - 0.5 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Gaji Pokok")
    c.drawString(x + 10 * cm, y, f": Rp. {item['basic']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Tunjangan Jabatan")
    c.drawString(x + 10 * cm, y, f": Rp. {item['jabatan']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Tunjangan Cuti")
    c.drawString(x + 10 * cm, y, f": Rp. {item['cuti']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Tanggal merah, Lembuh, Dll")
    c.drawString(x + 10 * cm, y, f": Rp. {(item['holiday'] + item['event']):,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Lembur Menggantikan")
    c.drawString(x + 10 * cm, y, f": Rp. {item['lembur']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Revisi, Kekurangan bulan lalu, dll")
    c.drawString(x + 10 * cm, y, f": Rp. {item['revisi']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "Tunjangan Hari Raya")
    c.drawString(x + 10 * cm, y, f": Rp. {item['thr']:,.2f}")
    y -= 0.5 * cm

    c.drawString(x, y, "PP 21 yang di bayar kantor")
    c.drawString(x + 10 * cm, y, f": Rp. {item['pph_kantor']:,.2f}")
    y -= 0.5 * cm

    c.setFont("Helvetica-Bold", 11)
    total_penerimaan = item['basic'] + item['jabatan'] + item['cuti'] + item['holiday'] + item['event']+ item['lembur'] + item['revisi'] + item['pph_kantor']
    c.drawString(x, y, "TOTAL PENERIMAAN")
    c.drawString(x + 10 * cm, y, f": Rp. {total_penerimaan:,.2f}")

    return y - 1 * cm

def render_potongan(c, item, x, start_y):
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, start_y, "POTONGAN")
    y = start_y - 0.5 * cm

    c.setFont("Helvetica", 11)

    bpjs_kes = item['potongan_bpjskes'] - item['tunjangan_bpjskes']
    c.drawString(x, y, "BPJS KESEHATAN")
    c.drawString(x + 10 * cm, y, f": Rp. {bpjs_kes:,.2f}")
    y -= 0.5 * cm

    bpjs_tk = item['potongan_bpjstk'] - item['tunjangan_bpjstk']
    c.drawString(x, y, "BPJS KESEHATAN")
    c.drawString(x + 10 * cm, y, f": Rp. {bpjs_tk:,.2f}")
    y -= 0.5 * cm
    
    c.drawString(x, y, "Ijin, Sakit, Kekurangan jam, dll ")
    c.drawString(x + 10 * cm, y, f": Rp. {item['kekurangan_jam']:,.2f}")
    y -= 0.5 * cm

    revisi = item['kelebihan'] + item['operasional']
    c.drawString(x, y, "Revisi, Kelebihan Bulan lalu, Operasional, dll")
    c.drawString(x + 10 * cm, y, f": Rp. {revisi:,.2f}")
    y -= 0.5 * cm
    
    c.drawString(x, y, "Pajak PPH 21")
    c.drawString(x + 10 * cm, y, f": Rp. {item['pph21']:,.2f}")
    y -= 0.5 * cm
    
    c.setFont("Helvetica-Bold", 11)
    total_potongan = bpjs_kes + bpjs_tk + item['kekurangan_jam'] + revisi + item['pph21']
    c.drawString(x, y, "TOTAL POTONGAN")
    c.drawString(x + 10 * cm, y, f": Rp. {total_potongan:,.2f}")

    return y - 1* cm

def render_footer(c, item, margin_x, start_y):
    pass

def multi_slip(data, filename="data/slip_gaji1.pdf"):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    for item in data:
        margin_x = 2 * cm
        start_y = height - 5 * cm   # Inisialisasi bari bertama, kecuali judul
        start_y = render_header(c, item, margin_x, start_y)
        start_y = render_penerimaan(c, item, margin_x, start_y)
        start_y = render_potongan(c, item, margin_x, start_y)
        c.showPage()
    c.save()

multi_slip(data)

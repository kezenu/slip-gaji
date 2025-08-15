import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from datetime import date
from decimal import Decimal, ROUND_DOWN


bulan_indonesia = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

# Baca data
df = pd.read_excel('gaji_contoh.xlsx')

# Mengisi semua data yang kosong dengan nilai 0
df.fillna(0, inplace=True)

df['nik'] = df['nik'].apply(lambda x: str(int(x)) if isinstance(x,(int, float)) else str(x))

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

def render_kop(c):
    c.drawImage("logo1.jpg", # Path file gambar
                3 * cm, # Koordinat x
                height - 2.8 * cm, # Koordinat y
                width= 2.3 * cm, # lebar gambar dalam cm
                height= 2.3 * cm, # Tinggi gambar dalam cm
                preserveAspectRatio=True,
                mask="auto"
    )
    x = 6 * cm
    y = height - 1 * cm
    c.setFont("Times-Roman", 16)
    c.setFillColor(colors.red)
    c.drawString(x, y, "PT. ALKATRA")
    c.setFillColor(colors.black)
    y -= 0.5 * cm

    c.setFont("Times-Roman", 12)
    c.drawString(x, y, "Komplek Ruko Puri Mestika, Blok I Nomor 6,  Batam Center, Batam.")
    y -= 0.5 * cm

    c.setFont("Times-Roman", 12)
    c.drawString(x, y, "Telp. : 0778-4806-115 ; 0822-6848-0475")
    y -= 0.5 * cm

    c.setFont("Times-Roman", 12)
    c.drawString(x, y, "Email : pt.alkatra@gmail.com; Website : https://pt-alkatra.business.site")
    y -= 0.5 * cm
    
    # Garis dibawah kop surat
    c.setLineWidth(1)
    c.line( 1 * cm, y, width - 1 * cm, y)
    c.setLineWidth(2)
    c.line( 1 * cm, y - 0.1 * cm, width - 1 * cm, y - 0.1 * cm)

def render_header(c, item, x, start_y):

    text_width = c.stringWidth("SLIP GAJI KARYAWAN", "Helvetica-Bold", 16)
    center = (width - text_width) / 2
    c.setFont("Helvetica-Bold", 16)
    c.drawString(center, start_y, "SLIP GAJI KARYAWAN")

    # Underline SLIP GAJI KARYAWAN
    c.setLineWidth(3)
    c.line(center , start_y - 0.3 * cm, center + text_width, start_y - 0.3 * cm)
    y = start_y - 2 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Nama")
    c.drawString(x + 5 * cm, y, f": {item['nama']}")
    y -= 0.5 * cm
    
    c.drawString(x, y, "NIK")
    c.drawString(x + 5 * cm, y, f": {item['nik']}")
    y -= 0.5 * cm

    c.drawString(x, y, "Lokasi")
    c.drawString(x + 5 * cm, y, f": {item['lokasi']}")
    y -= 0.5 * cm

    c.drawString(x, y, "Status")
    c.drawString(x + 5 * cm, y, f": {item['status']}")
    y -= 0.5 * cm
    
    c.drawString(x, y, "Periode Gaji")
    c.drawString(x + 5 * cm, y, f": {bulan_indonesia[item['bulan'].month]} {item['bulan'].year}")
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

    c.drawString(x, y, "Tanggal merah, Lembur, dll")
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
    item['total_penerimaan'] = item['basic'] + item['jabatan'] + item['cuti'] + item['holiday'] + item['event']+ item['lembur'] + item['revisi'] + item['pph_kantor']
    c.drawString(x, y, "TOTAL PENERIMAAN")
    c.drawString(x + 10 * cm, y, f": Rp. {item['total_penerimaan']:,.2f}")

    # Deckorasi pada bagian penerimaan
    c.setLineWidth(1)
    c.roundRect(x - 0.3 * cm, y - 0.3 * cm , 15.5 * cm, 5.5 * cm, radius=20, stroke=1, fill=0)

    return y - 1.5 * cm

def render_potongan(c, item, x, start_y):
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, start_y, "POTONGAN")
    y = start_y - 0.5 * cm

    c.setFont("Helvetica", 11)

    bpjs_kes = item['potongan_bpjskes'] - item['tunjangan_bpjskes']
    c.drawString(x, y, "BPJS Kesehatan")
    c.drawString(x + 10 * cm, y, f": Rp. {bpjs_kes:,.2f}")
    y -= 0.5 * cm

    bpjs_tk = item['potongan_bpjstk'] - item['tunjangan_bpjstk']
    c.drawString(x, y, "BPJS Ketenagakerjaan")
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
    item['total_potongan'] = bpjs_kes + bpjs_tk + item['kekurangan_jam'] + revisi + item['pph21']
    c.drawString(x, y, "TOTAL POTONGAN")
    c.drawString(x + 10 * cm, y, f": Rp. {item['total_potongan']:,.2f}")

    return y - 1* cm

def render_footer(c, item, x, start_y):
    c.setFont("Helvetica-Bold", 12)
    item['total_diterima'] = item['total_penerimaan'] - item['total_potongan']
    c.drawString(x, start_y, "TOTAL DITERIMA ")
    angka = Decimal(str(item['total_diterima'])).quantize(Decimal("0.00"), rounding=ROUND_DOWN)
    c.drawString(x + 10 * cm, start_y, f": Rp. {angka:,.2f}")
    y = start_y - 2 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Mengetahui")
    y -= y - 6 * cm
    c.drawString(x, y, "Joel Simatupang")
    y -= y - 6.5 * cm
    c.drawString(x, y, "HRD")
    y -= y - 3.5 * cm

    c.setFont("Helvetica", 9)
    hari_ini = date.today()
    c.drawString(1 * cm, 2 * cm, f"Dokumen ini di cetak otomatis oleh sistem pada tanggal {hari_ini.strftime('%d-%m-%Y')}")


def multi_slip(data, output_dir="data/slip"):
    os.makedirs(output_dir, exist_ok=True)

    for item in data:
        nama = item['nama'].replace(" ", "_")
        nik = str(item['nik'])
        filepath = os.path.join(output_dir, f"{nik}_{nama}.pdf") 
        c = canvas.Canvas(filepath, pagesize=A4)
        width, height = A4

        render_kop(c)
        margin_x = 2 * cm
        start_y = height - 4 * cm   # Inisialisasi bari bertama, kecuali judul
        start_y = render_header(c, item, margin_x, start_y)
        start_y = render_penerimaan(c, item, margin_x, start_y)
        start_y = render_potongan(c, item, margin_x, start_y)
        start_y = render_footer(c, item, margin_x, start_y)
        c.showPage()
        c.save()

multi_slip(data)
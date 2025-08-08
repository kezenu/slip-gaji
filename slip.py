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

def multi_slip(data, filename="data/slip_gaji.pdf"):
    c = canvas.Canvas(filename, pagesize= A4)
    width, height = A4

    for item in data:
        c.setFont("Helvetica-Bold", 16)
        c.drawString(8 * cm, height - 5 * cm, "SLIP GAJI KARYAWAN")

        #  Informasi pekerja
        c.setFont("Helvetica", 11)
        c.drawString(1 * cm, height - 6.5 * cm, f"Nama")
        c.drawString(1 * cm, height - 7 * cm, f"NIK")
        c.drawString(1 * cm, height - 7.5 * cm, f"Lokasi")
        c.drawString(1 * cm, height - 8 * cm, f"Status")
        c.drawString(1 * cm, height - 8.5 * cm, f"Periode Gaji")

        # 
        c.drawString(5 * cm, height - 6.5 * cm, f" : {item['nama']}")
        c.drawString(5 * cm, height - 7 * cm, f" : {item['nik']:.0f}")
        c.drawString(5 * cm, height - 7.5 * cm, f" : {item['lokasi']}")
        c.drawString(5 * cm, height - 8 * cm, f" : {item['status']}")
        c.drawString(5 * cm, height - 8.5 * cm, f" : {bulan_indonesia[item['bulan'].month]}")

        # RINCIAN PENERIMAAN
        c.setFont("Helvetica-Bold", 12)
        c.drawString(1 * cm, height - 10 * cm, f"PENERIMAAN")
        c.setFont("Helvetica", 11)

        #  Indek
        c.drawString(1 * cm, height - 11 * cm, f"Gaji Pokok")
        c.drawString(1 * cm, height - 11.5 * cm, f"Tunjangan Jabatan")
        c.drawString(1 * cm, height - 12 * cm, f"Tunjangan Cuti")
        c.drawString(1 * cm, height - 12.5 * cm, f"Tanggal Merah, Lembur, dll")
        c.drawString(1 * cm, height - 13 * cm, f"Lembur Menggantikan")
        c.drawString(1 * cm, height - 13.5 * cm, f"Tunjangan Hari Raya")
        c.drawString(1 * cm, height - 14 * cm, f"PPH 21 dibayar kantor")
        c.drawString(1 * cm, height - 14.5 * cm, f"Revisi, Kurang bulan lalu, dll")
        c.drawString(1 * cm, height - 15.5 * cm, f"TOTAL PENERIMAAN")

        # Nominal
        c.drawString(8 * cm, height - 11 * cm, f" : Rp. {item['basic']:,.2f}")
        c.drawString(8 * cm, height - 11.5 * cm, f" : Rp. {item['jabatan']:,.2f}")
        c.drawString(8 * cm, height - 12 * cm, f" : Rp. {item['cuti']:,.2f}")
        c.drawString(8 * cm, height - 12.5 * cm, f" : Rp. {(item['holiday'] + item['event']):,.2f}")
        c.drawString(8 * cm, height - 13 * cm, f" : Rp. {item['lembur']:,.2f}")
        c.drawString(8 * cm, height - 13.5 * cm, f" : Rp. {item['thr']:,.2f}")
        c.drawString(8 * cm, height - 14 * cm, f" : Rp. {item['pph_kantor']:,.2f}")
        c.drawString(8 * cm, height - 14.5 * cm, f" : Rp. {item['revisi']:,.2f}")
        total = item['basic'] + item['jabatan'] + item['cuti'] + item['holiday'] + item['event']+ item['lembur'] + item['revisi']
        c.setFont("Helvetica-Bold", 12)
        c.drawString(8 * cm, height - 15.5 * cm, f" : Rp. {total:,.2f}")
        

        # RINCIAN POTONGAN
        c.setFont("Helvetica-Bold", 12)
        c.drawString(1 * cm, height - 17 * cm, f"POTONGAN")
        c.setFont("Helvetica", 11)

        # Indek
        c.drawString(1 * cm, height - 17.5 * cm, f"Potongan BPJSTK")
        c.drawString(1 * cm, height - 18 * cm, f"Potongan BPJSKES")
        c.drawString(1 * cm, height - 18.5 * cm, f"Absen, ijin, kurang jam/hari, dll")

        # Nominal
        bpjstk = item['potongan_bpjstk'] - item['tunjangan_bpjstk']
        bpjskes = item['potongan_bpjskes'] - item['tunjangan_bpjskes']
        c.drawString(8 * cm, height - 17.5 * cm, f" : Rp. {bpjstk:,.2f}")
        c.drawString(8 * cm, height - 18 * cm, f" : Rp. {bpjskes:,.2f}")
        c.drawString(8 * cm, height - 18.5 * cm, f" : Rp. {item['kekurangan_jam']:,.2f}")
        # ⛳️ Halaman baru untuk setiap slip
        c.showPage()

    c.save()

# Jalankan fungsi
multi_slip(data)

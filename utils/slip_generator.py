import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from datetime import date
from decimal import Decimal, ROUND_DOWN
from PyPDF2 import PdfReader, PdfWriter
from indonesian_number_normalizer import create_normalizer


# =========================
#  KONFIGURASI
# =========================
EXCEL_PATH = "assets/api/gaji_contoh.xlsx"
LOGO_PATH = "assets/logo1.jpg"
OUTPUT_DIR = "data/slip"

bulan_indonesia = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]


# =========================
#  DATA CONTROLLER
# =========================
def controller_data():
    df = pd.read_excel(EXCEL_PATH)
    df.fillna(0, inplace=True)
    df["nik"] = df["nik"].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x))
    df.set_index("nik", inplace=True)

    # pastikan numeric
    numeric_cols = [
        "basic", "jabatan", "tunjangan_bpjstk", "tunjangan_bpjskes", "cuti", "holiday",
        "lembur", "event", "total_ot", "revisi", "thr", "pph_kantor", "bruto", "operasional",
        "kelebihan", "kekurangan_jam", "potongan_bpjstk", "pph21", "total_potongan", "netto"
    ]
    df[numeric_cols] = df[numeric_cols].astype(int)

    return df.reset_index().to_dict(orient="records")


# =========================
#  HELPER
# =========================
def draw_field(c, label, value, x, y, offset=12 * cm):
    """Helper untuk menggambar field label : value"""
    c.drawString(x, y, label)
    c.drawString(x + offset, y, f": {value}")


# =========================
#  RENDERING
# =========================
def render_kop(c, width, height):
    c.drawImage(LOGO_PATH, 3 * cm, height - 2.8 * cm, 2.3 * cm, 2.3 * cm, preserveAspectRatio=True, mask="auto")
    x, y = 6 * cm, height - 1 * cm

    c.setFont("Times-Roman", 16)
    c.setFillColor(colors.red)
    c.drawString(x, y, "PT. ALKATRA")

    c.setFillColor(colors.black)
    y -= 0.5 * cm
    c.setFont("Times-Roman", 12)
    c.drawString(x, y, "Komplek Ruko Puri Mestika, Blok I Nomor 6, Batam Center, Batam.")
    y -= 0.5 * cm
    c.drawString(x, y, "Telp. : 0778-4806-115 ; 0822-6848-0475")
    y -= 0.5 * cm
    c.drawString(x, y, "Email : pt.alkatra@gmail.com; Website : https://pt-alkatra.business.site")
    y -= 0.5 * cm

    # garis
    c.setLineWidth(1)
    c.line(1 * cm, y, width - 1 * cm, y)
    c.setLineWidth(2)
    c.line(1 * cm, y - 0.1 * cm, width - 1 * cm, y - 0.1 * cm)


def render_header(c, item, x, y, width):
    text = "SLIP GAJI KARYAWAN"
    text_width = c.stringWidth(text, "Helvetica-Bold", 16)
    c.setFont("Helvetica-Bold", 16)
    c.drawString((width - text_width) / 2, y, text)

    c.setLineWidth(3)
    c.line((width - text_width) / 2, y - 0.3 * cm, (width + text_width) / 2, y - 0.3 * cm)

    c.setFont("Helvetica", 11)
    y -= 2 * cm
    offset = 5 * cm 
    draw_field(c, "Nama", item["nama"], x, y, offset); y -= 0.5 * cm
    draw_field(c, "NIK", item["nik"], x, y, offset); y -= 0.5 * cm
    draw_field(c, "Lokasi", item["lokasi"], x, y, offset); y -= 0.5 * cm
    draw_field(c, "Status", item["status"], x, y, offset); y -= 0.5 * cm
    draw_field(c, "Periode Gaji", f"{bulan_indonesia[item['bulan'].month]} {item['bulan'].year}", x, y, offset)
    return y - 1 * cm


def render_penerimaan(c, item, x, y, width):
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, "PENERIMAAN"); y -= 0.5 * cm
    c.setFont("Helvetica", 11)

    fields = [
        ("Gaji Pokok", item["basic"]),
        ("Tunjangan Jabatan", item["jabatan"]),
        ("Tunjangan Cuti", item["cuti"]),
        ("Tanggal merah, Lembur, dll", item["holiday"] + item["event"]),
        ("Lembur Menggantikan", item["lembur"]),
        ("Revisi, Kekurangan bulan lalu, dll", item["revisi"]),
        ("Tunjangan Hari Raya", item["thr"]),
        ("PP 21 yang di bayar kantor", item["pph_kantor"])
    ]

    for label, val in fields:
        draw_field(c, label, f"Rp. {val:,.2f}", x, y); y -= 0.5 * cm

    c.setFont("Helvetica-Bold", 11)
    draw_field(c, "TOTAL PENERIMAAN", f"Rp. {item['total_penerimaan']:,.2f}", x, y)

    # dekorasi
    c.setLineWidth(1)
    c.roundRect(1.5 * cm, y - 0.3 * cm, width - 3 * cm, 5.5 * cm, radius=20, stroke=1, fill=0)
    return y - 1.5 * cm


def render_potongan(c, item, x, y, width):
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, "POTONGAN"); y -= 0.5 * cm
    c.setFont("Helvetica", 11)

    bpjs_kes = item["potongan_bpjskes"] - item["tunjangan_bpjskes"]
    bpjs_tk = item["potongan_bpjstk"] - item["tunjangan_bpjstk"]
    revisi = item["kelebihan"] + item["operasional"]

    fields = [
        ("BPJS Kesehatan", bpjs_kes),
        ("BPJS Ketenagakerjaan", bpjs_tk),
        ("Ijin, Sakit, Kekurangan jam, dll", item["kekurangan_jam"]),
        ("Revisi, Kelebihan Bulan lalu, Operasional, dll", revisi),
        ("Pajak PPH 21", item["pph21"])
    ]
    for label, val in fields:
        draw_field(c, label, f"Rp. {val:,.2f}", x, y); y -= 0.5 * cm

    c.setFont("Helvetica-Bold", 11)
    draw_field(c, "TOTAL POTONGAN", f"Rp. {item['total_potongan']:,.2f}", x, y)

    c.setLineWidth(1)
    c.roundRect(1.5 * cm, y - 0.3 * cm, width - 3 * cm, 4 * cm, radius=20, stroke=1, fill=0)
    return y - 1 * cm


def render_footer(c, item, x, y, width):
    a = 12 * cm
    angka = Decimal(str(item["total_diterima"])).quantize(Decimal("0.00"), rounding=ROUND_DOWN)

    c.setFont("Helvetica-Bold", 12)
    y -= 0.5 * cm
    c.drawString(x, y, "TOTAL DITERIMA ")
    c.drawString(x + a, y, f": Rp. {angka:,.2f}")

    c.setLineWidth(1)
    c.roundRect(1.5 * cm, y - 1.3 * cm, width - 3 * cm, 2 * cm, radius=10, stroke=1, fill=0)
    y -= 1 * cm

    normalize = create_normalizer()
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, f"{normalize.number_to_words(item['total_diterima']).title()} Rupiah")
    y -= 2 * cm

    c.setFont("Helvetica", 11)
    c.drawString(x, y, "Mengetahui"); y -= 2 * cm
    c.drawString(x, y, "Joel Simatupang"); y -= 0.5 * cm
    c.drawString(x, y, "HRD")

    c.setFont("Helvetica", 9)
    hari_ini = date.today()
    c.drawString(1 * cm, 2 * cm, f"Dokumen ini dicetak otomatis pada {hari_ini.strftime('%d-%m-%Y')}")


# =========================
#  SLIP
# =========================
def generate_slip_pdf(item, output_path):
    width, height = A4
    c = canvas.Canvas(output_path, pagesize=A4)

    # hitung total sebelum render
    item["total_penerimaan"] = item["basic"] + item["jabatan"] + item["cuti"] + item["holiday"] + item["event"] + item["lembur"] + item["revisi"] + item["pph_kantor"]
    bpjs_kes = item["potongan_bpjskes"] - item["tunjangan_bpjskes"]
    bpjs_tk = item["potongan_bpjstk"] - item["tunjangan_bpjstk"]
    revisi = item["kelebihan"] + item["operasional"]
    item["total_potongan"] = bpjs_kes + bpjs_tk + item["kekurangan_jam"] + revisi + item["pph21"]
    item["total_diterima"] = item["total_penerimaan"] - item["total_potongan"]

    render_kop(c, width, height)
    y = render_header(c, item, 2 * cm, height - 4 * cm, width)
    y = render_penerimaan(c, item, 2 * cm, y, width)
    y = render_potongan(c, item, 2 * cm, y, width)
    render_footer(c, item, 2 * cm, y, width)

    c.showPage()
    c.save()


def protect_pdf(input_path, output_path, password):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    if password:
        writer.encrypt(password)
    with open(output_path, "wb") as f:
        writer.write(f)


def multi_slip(data, output_dir=OUTPUT_DIR):
    os.makedirs(output_dir, exist_ok=True)
    for item in data:
        nama_file = item["nama"].replace(" ", "_")
        temp_path = os.path.join(output_dir, f"_temp_{nama_file}.pdf")
        final_path = os.path.join(output_dir, f"{nama_file}.pdf")

        generate_slip_pdf(item, temp_path)
        protect_pdf(temp_path, final_path, str(int(item["nik"])))
        os.remove(temp_path)
        print(f"✅ Slip {item['nama']} selesai → {final_path}")


def single_slip(data, identifier, output_dir=OUTPUT_DIR):
    os.makedirs(output_dir, exist_ok=True)
    for item in data:
        if str(item["nik"]) == str(identifier) or item["nama"].lower() == str(identifier).lower():
            nama_file = item["nama"].replace(" ", "_")
            temp_path = os.path.join(output_dir, f"_temp_{nama_file}.pdf")
            final_path = os.path.join(output_dir, f"{nama_file}.pdf")

            generate_slip_pdf(item, temp_path)
            protect_pdf(temp_path, final_path, None)  # tanpa password
            os.remove(temp_path)
            print(f"✅ Slip {item['nama']} selesai → {final_path}")
            return
    print("❌ Data tidak ditemukan.")


# testing
nama = "ramli"
single_slip(controller_data(), nama)

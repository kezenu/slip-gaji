from utils.slip_generator import multi_slip, controller_data
from utils.sent_email import kirim_email

while True:
    print("1. Generate Slip Gaji")
    print("2. kirim slip gaji dengan email")
    print("0. Keluar")

    pilihan = int(input("Masukan pilihan : "))
    if pilihan == 1:
        multi_slip(controller_data())
        print("Slip gaji berhasil dicetak")
        continue
    elif pilihan == 2:
        kirim_email(controller_data())
        continue
    else:
        break

from utils.slip_generator import multi_slip, controller_data
from utils.sent_email import kirim_email

while True:
    pilihan = int(input("Masukan pilihan : "))
    if pilihan == 1:
        print("Slip gaji berhasil dicetak")
        multi_slip(controller_data())
        continue
    elif pilihan == 2:
        kirim_email()
        continue
    else:
        break
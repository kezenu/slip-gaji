from utils.slip_generate import multi_slip, controller_data

while True:
    pilihan = int(input("Masukan pilihan : "))
    if pilihan == 1:
        print("Slip gaji berhasil dicetak")
        multi_slip(controller_data())
        continue
    elif pilihan == 0:
        break
    else:
        break

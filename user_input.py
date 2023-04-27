from openpyxl import load_workbook
import pyinputplus as pyip

#MENDAPATKAN DATA WP DARI EXCEL
wb = load_workbook("./excel/wpairtanah.xlsx")
ws = wb.active

daftar_wp = {}

for i in ws['A']:
    if(i.row == 1):
        continue

    daftar_wp[i.offset(column = 1).value] = i.value

#INPUT DATA
wajib_pajak = pyip.inputMenu([i for i in daftar_wp], numbered = True)
npwpd = daftar_wp[wajib_pajak]
print(wajib_pajak, npwpd, sep = " - ")
seleksi_mode = pyip.inputMenu(["Bulan", "Bulan berurutan", "Bulan tidak berurut"], numbered = True)
list_bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
masa_pajak = []

#INPUT PROCESSING
if(seleksi_mode == "Bulan"):
    masa_pajak.append(pyip.inputMenu(list_bulan, numbered = True, prompt = "Masa pajak: \n"))
    print(masa_pajak)
elif(seleksi_mode == "Bulan berurutan"):
    masa_awal = list_bulan.index(pyip.inputMenu(list_bulan, numbered = True, prompt = "Masa pajak awal: \n"))
    masa_akhir = list_bulan.index(pyip.inputMenu(list_bulan, numbered = True, prompt = "Masa pajak akhir: \n"))

    for i in range(masa_awal, masa_akhir + 1):
        masa_pajak.append(list_bulan[i])
    
    print(masa_pajak)
else:
    banyak_bulan = int(input("Berapa banyak bulan? "))

    for i in range(banyak_bulan):
        masa_pajak.append(pyip.inputMenu(list_bulan, numbered = True, prompt = "Masa pajak: \n"))
    
    print(masa_pajak)
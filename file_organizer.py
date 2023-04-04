from pdfminer.high_level import extract_pages, extract_text
from datetime import datetime
import pyinputplus as pyip
import re
import os
import shutil

def extractData(file):
    text = extract_text(file)
    translate_bulan = {"January": "Januari", "February": "Februari", "March": "Maret", "April": "April", "May": "Mei", "June": "Juni", "July": "Juli", "August": "Agustus", "September": "September", "October": "Oktober",
                    "November": "November", "December": "Desember"}
    
    #REGEX
    identitas_wp_regex = re.compile(r":\n+(.+)\n(.+|.+\n.+)\n+(.+)\n+P.\d.\d{7}.\d{2}.\d{2}")
    jenis_pajak_regex = re.compile(r"SURAT KETETAPAN PAJAK DAERAH\n(.+)\n")
    nomor_bayar_regex = re.compile(r"NO. BAYAR :\n(.+)\n")
    nomor_kohir_regex = re.compile(r"NO. KOHIR :\n(.+)\n")
    jatuh_tempo_regex = re.compile(r"P.\d.\d{7}.\d{2}.\d{2}\n+(.+)\n")
    npwpd_regex = re.compile(r"(P.\d.\d{7}.\d{2}.\d{2})")
    volume_regex = re.compile(r"Volume : (.+ M3)")
    periode_regex = re.compile(r"Periode : (\d{2}-\d{2}-\d{4}) s/d (\d{2}-\d{2}-\d{4})")
    jumlah_ketetapan_regex = re.compile(r"Jumlah Ketetapan Pokok Pajak\n+(.+)")

    nama_wp = identitas_wp_regex.findall(text)[0][0]
    alamat_wp = identitas_wp_regex.findall(text)[0][1]
    kecamatan_wp = identitas_wp_regex.findall(text)[0][2]
    jenis_pajak = jenis_pajak_regex.findall(text)[0]
    nomor_bayar = nomor_bayar_regex.findall(text)[0]
    nomor_kohir = nomor_kohir_regex.findall(text)[0]
    jatuh_tempo = jatuh_tempo_regex.findall(text)[0]
    npwpd = npwpd_regex.findall(text)[0]
    volume = volume_regex.findall(text)[0]
    periode_awal = periode_regex.findall(text)[0][0]
    periode_akhir = periode_regex.findall(text)[0][1]
    jumlah_ketetapan = jumlah_ketetapan_regex.findall(text)[0]

    #masa_pajak
    temp = datetime.strptime(periode_awal, "%d-%m-%Y")
    masa_pajak_temp = datetime.strftime(temp, "%B")
    masa_pajak = translate_bulan[masa_pajak_temp].upper()
    #tahun_pajak
    tahun_pajak = datetime.strftime(temp, "%Y")

    #EXCEPTION IN FILE NAMING

    if(nama_wp == "PT.SU INDONESIA  DH."):
        nama_wp = "PT.SU INDONESIA"

    print(nama_wp)

    return {"nama wp": nama_wp, "alamat": alamat_wp, "kecamatan": kecamatan_wp, "masa pajak": masa_pajak, "tahun pajak": tahun_pajak, "jenis pajak": jenis_pajak, "nomor bayar": nomor_bayar, "nomor kohir": nomor_kohir, 
            "jatuh tempo": jatuh_tempo, "npwpd": npwpd, "volume": volume, "periode awal" : periode_awal, "periode akhir": periode_akhir, "jumlah ketetapan": jumlah_ketetapan, "dokumen": "SKPD"}
    
def renameFile(file, masa_pajak, tahun_pajak, nama_wp, tipe_dokumen):
    filename = "%s (%s %s) %s" %(tipe_dokumen, masa_pajak, tahun_pajak, nama_wp)

    os.rename(file, "%s.pdf" %(filename))

    return filename

def organizeFiles():
    os.chdir("raw_documents")
    filenames = sorted(os.listdir())

    for i in filenames:
        data = extractData(i)
        renamed = renameFile(i, data["masa pajak"], data["tahun pajak"], data["nama wp"], data["dokumen"])

        moveFile("%s.pdf" %renamed, data["nama wp"])

    os.chdir("..")

def moveFile(file, nama_wp):
    destination_path = "../wp/%s" %nama_wp

    if (not os.path.exists(destination_path)):
        os.mkdir("../wp/%s" %nama_wp)
    
    os.rename(file, "%s/%s" %(destination_path, file))
    

organizeFiles()
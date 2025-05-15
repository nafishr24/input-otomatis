from selenium import webdriver
from openpyxl import load_workbook
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

start_time = time.time()

wb = load_workbook(filename="Peringkat SKD.xlsx")
sheet_range = wb['Sheet2']
driver = webdriver.Chrome()
driver.get("https://docs.google.com/forms/d/e/1FAIpQLSeCCnU4_cu3CxXLgqavXO3FpvszBPgOAa1_Di-CI3pOTAGo1g/viewform")
driver.maximize_window()
driver.implicitly_wait(1)

i = 2
while i <= len(sheet_range['A']):
    # Ambil data dari Excel
    peringkat = sheet_range['A' + str(i)].value
    no_peserta = sheet_range['B' + str(i)].value
    nama = sheet_range['C' + str(i)].value
    manajerial = sheet_range['D' + str(i)].value
    sosio = sheet_range['E' + str(i)].value
    teknis = sheet_range['F' + str(i)].value
    wawancara = sheet_range['G' + str(i)].value
    total = sheet_range['H' + str(i)].value
    posisi = sheet_range['I' + str(i)].value

    # Isi form
    fields = [
        peringkat, no_peserta, nama, manajerial,
        sosio, teknis, wawancara, total, posisi
    ]
    
    for n, value in enumerate(fields, start=1):
        xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[{n}]/div/div/div[2]/div/div[1]/div/div[1]/input'
        driver.find_element('xpath', xpath).send_keys(str(value))
    
    # Submit
    driver.find_element('xpath','//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()
    
    # Tunggu sampai form baru siap
    time.sleep(1)
    
    # Klik link untuk mengisi form lagi jika ada
    try:
        driver.find_element('link text', 'Submit another response').click()
    except:
        # Jika tidak ada, refresh halaman
        driver.get("https://docs.google.com/forms/d/e/1FAIpQLSeCCnU4_cu3CxXLgqavXO3FpvszBPgOAa1_Di-CI3pOTAGo1g/viewform")
    
    time.sleep(1)
    i += 1

print("Telah selesai")

# Catat waktu selesai dan hitung durasi

end_time = time.time()
total_time = end_time - start_time

# Konversi ke menit dan detik jika lebih dari 60 detik
if total_time > 60:
    minutes = int(total_time // 60)
    seconds = int(total_time % 60)
    print(f"Proses selesai! Total waktu: {minutes} menit {seconds} detik")
else:
    print(f"Proses selesai! Total waktu: {total_time:.2f} detik")

driver.quit()
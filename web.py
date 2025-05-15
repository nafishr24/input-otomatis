from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import time

# Catat waktu mulai
start_time = time.time()

# Inisialisasi WebDriver
driver = webdriver.Chrome()
driver.implicitly_wait(10)  # Waktu tunggu implisit

# Langkah 1: Buka halaman login
login_url = "https://example.com/login"  # Ganti dengan URL login
driver.get(login_url)

# Langkah 2: Isi form login
username = "your_username"  # Ganti dengan username
password = "your_password"  # Ganti dengan password

driver.find_element(By.ID, "username").send_keys(username)  # Ganti ID sesuai elemen
driver.find_element(By.ID, "password").send_keys(password)
driver.find_element(By.XPATH, "//button[@type='submit']").click()  # Ganti XPATH tombol login

# Langkah 3: Tunggu sampai login sukses (misal: muncul elemen dashboard)
try:
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//h1[contains(text(), 'Dashboard')]"))
    )
    print("Login berhasil!")
except Exception as e:
    print("Gagal login:", e)
    driver.quit()
    exit()

# Langkah 4: Buka halaman input data (URL berbeda)
input_url = "https://example.com/input-form"  # Ganti dengan URL form input
driver.get(input_url)

# Langkah 5: Baca data Excel dan isi form
wb = load_workbook(filename="Peringkat SKD.xlsx")
sheet_range = wb['Sheet2']

i = 2
while i <= len(sheet_range['A']):
    # Langkah 5a: Klik tombol "Add" untuk membuka form baru
    try:
        add_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add')]"))
        )
        add_button.click()
        print("Tombol Add diklik.")
    except Exception as e:
        print("Gagal menemukan/mengklik tombol Add:", e)
        driver.quit()
        exit()

    # Tunggu sampai form input muncul
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//h3[contains(text(), 'Form Baru')]"))
        )
    except Exception as e:
        print("Form input tidak muncul setelah Add:", e)
        driver.quit()
        exit()

    # Langkah 5b: Ambil data dari Excel
    peringkat = sheet_range['A' + str(i)].value
    no_peserta = sheet_range['B' + str(i)].value
    nama = sheet_range['C' + str(i)].value
    manajerial = sheet_range['D' + str(i)].value
    sosio = sheet_range['E' + str(i)].value
    teknis = sheet_range['F' + str(i)].value
    wawancara = sheet_range['G' + str(i)].value
    total = sheet_range['H' + str(i)].value
    posisi = sheet_range['I' + str(i)].value

    # Langkah 5c: Isi form (ini untuk google form, sesuaikan dengan yang di web)
    fields = [peringkat, no_peserta, nama, manajerial, sosio, teknis, wawancara, total, posisi]
    for n, value in enumerate(fields, start=1):
        xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[{n}]/div/div/div[2]/div/div[1]/div/div[1]/input'
        driver.find_element(By.XPATH, xpath).send_keys(str(value))
    
    # Langkah 5d: Submit form
    driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()
    time.sleep(2)  # Tunggu setelah submit

    # Langkah 5e: Kembali ke halaman input (jika diperlukan)
    driver.get(input_url)
    time.sleep(1)
    i += 1

# Hitung total waktu
end_time = time.time()
total_time = end_time - start_time
if total_time > 60:
    print(f"Proses selesai! Total waktu: {int(total_time // 60)} menit {int(total_time % 60)} detik")
else:
    print(f"Proses selesai! Total waktu: {total_time:.2f} detik")

driver.quit()
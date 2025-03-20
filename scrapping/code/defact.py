from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os

# Setup Selenium WebDriver
options = Options()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Cek apakah file Excel hasil scraping sebelumnya ada
file_excel = "hasil_scraping.xlsx"
if not os.path.exists(file_excel):
    print("❌ File hasil_scraping.xlsx tidak ditemukan!")
    exit()

# Baca file Excel yang sudah ada
xls = pd.ExcelFile(file_excel)

# Ambil ID dari sheet "Arabica" dan "Robusta"
ids = []
for sheet_name in ["Arabica", "Robusta"]:
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if "ID" in df.columns:
            # Hapus tanda '#' di depan ID
            cleaned_ids = df["ID"].astype(str).str.replace("#", "").tolist()
            ids.extend(cleaned_ids)
        else:
            print(f"⚠️ Kolom ID tidak ditemukan di sheet {sheet_name}!")
    else:
        print(f"⚠️ Sheet {sheet_name} tidak ditemukan!")

# Batasi hanya 2 ID untuk testing
# ids = ids[:5]

if not ids:
    print("❌ Tidak ada ID yang ditemukan untuk scraping defect information!")
    driver.quit()
    exit()

data_defect = []

# Scraping Defect Information per ID
for id_text in ids:
    url = f"https://database.coffeeinstitute.org/coffee/{id_text}/green"
    driver.get(url)
    time.sleep(5)

    defect_data = {"ID": id_text}

    try:
        tables = driver.find_elements(By.CLASS_NAME, "grade_details")
        tables = tables[:2]  # Ambil dua tabel pertama saja
        
        if not tables:
            print(f"⚠️ Data defect untuk ID {id_text} tidak ditemukan!")
            continue
        # Ambil data dari tabel pertama
        rows = tables[0].find_elements(By.TAG_NAME, "tr")
        headers = [th.text.strip() for th in rows[0].find_elements(By.TAG_NAME, "th")][1:]
        values = [td.text.strip() for td in rows[1].find_elements(By.TAG_NAME, "td")]
        defect_data.update(dict(zip(headers, values)))
        
        # Ambil data dari tabel ke 2
        for i, table in enumerate(tables):
            rows = table.find_elements(By.TAG_NAME, "tr")

            if len(rows) < 4:
                print(f"⚠️ Tidak cukup data dalam tabel untuk ID {id_text}!")
                continue

            # Ambil header pertama (baris ke-2) dan data pertama (baris ke-3)
            headers_1 = [th.text.strip() for th in rows[0].find_elements(By.TAG_NAME, "th") if th.text.strip()]
            values_1 = [td.text.strip() for td in rows[1].find_elements(By.TAG_NAME, "td")]

            # Ambil header kedua (baris ke-4) dan data kedua (baris ke-5)
            headers_2 = [th.text.strip() for th in rows[2].find_elements(By.TAG_NAME, "th") if th.text.strip()]
            values_2 = [td.text.strip() for td in rows[3].find_elements(By.TAG_NAME, "td")]

            # Simpan ke dalam dictionary
            for header, value in zip(headers_1, values_1):
                defect_data[f"Tabel {i+1} - {header}"] = value

            for header, value in zip(headers_2, values_2):
                defect_data[f"Tabel {i+1} - {header}"] = value

        data_defect.append(defect_data)
        print(f"✅ Data defect ID {id_text} berhasil diambil!")
    except Exception as e:
        print(f"❌ Gagal mengambil data defect ID {id_text}: {e}")


# Simpan data defect ke sheet baru "Defect Information"
if data_defect:
    with pd.ExcelWriter(file_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_defect = pd.DataFrame(data_defect)
        df_defect.to_excel(writer, sheet_name="Defect Information", index=False)
    print("🎉 Semua data defect berhasil disimpan ke hasil_scraping.xlsx!")
else:
    print("⚠️ Tidak ada data defect yang berhasil di-scrape.")

# Tutup browser
driver.quit()

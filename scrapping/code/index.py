from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

# Setup Selenium WebDriver
options = Options()
# options.add_argument("--headless")  # Bisa diaktifkan jika tidak ingin membuka browser
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

# Inisialisasi WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# List URL yang ingin di-scrape
urls = {
    "Arabica": "https://database.coffeeinstitute.org/coffees/arabica",
    "Robusta": "https://database.coffeeinstitute.org/coffees/robusta"
}

# Tempat penyimpanan hasil scraping
data_main = {}  # Data Arabica & Robusta
data_detail = {
    "Sample Information": [],
    "Cupping Scores": [],
    "Green Analysis": [],
    "Certification Information": []
}

# Scraping Data Utama (Arabica & Robusta)
for sheet_name, url in urls.items():
    driver.get(url)
    # time.sleep(5)  # Tunggu halaman termuat
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "table")))


    all_rows = []  # Simpan semua data utama
    id_links = []  # Simpan ID & Link detail

    try:
        while True:
            # Cari tabel
            table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]

            # Ambil isi tabel
            rows = table.find_elements(By.TAG_NAME, "tr")[1:]  # Skip header
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = [cell.text.strip() for cell in cells]
                all_rows.append(row_data)

                # Cari ID dan link detail
                try:
                    id_link = row.find_element(By.TAG_NAME, "a")
                    id_text = id_link.text.strip().replace("#", "")
                    detail_url = id_link.get_attribute("href")
                    id_links.append((id_text, detail_url))
                except:
                    pass

            # Coba klik tombol "Next" jika ada
            try:
                next_button = driver.find_element(By.ID, "DataTables_Table_0_next")
                if "disabled" in next_button.get_attribute("class"):
                    break  # Berhenti jika tidak bisa lanjut
                next_button.click()
                time.sleep(3)  # Tunggu halaman berikutnya
            except:
                break

        # Simpan data utama ke dictionary
        data_main[sheet_name] = pd.DataFrame(all_rows, columns=headers)
        print(f"✅ Data utama dari {url} berhasil disimpan!")

    except Exception as e:
        print(f"❌ Gagal menemukan tabel di {url}: {e}")

    # Scraping Data Detail per ID
    for id_text, detail_url in id_links:
        try:
            driver.get(detail_url)
            time.sleep(5)  # Tunggu halaman detail termuat

            # Ambil Sample Information
            sample_info = {"ID": id_text}
            sample_table = driver.find_element(By.CLASS_NAME, "sample_information")
            rows = sample_table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                headers = row.find_elements(By.TAG_NAME, "th")
                if len(cells) == 2 and len(headers) == 2:  # Format tabel selalu 2 kolom header dan 2 kolom data
                    sample_info[headers[0].text.strip()] = cells[0].text.strip()
                    sample_info[headers[1].text.strip()] = cells[1].text.strip()
            data_detail["Sample Information"].append(sample_info)

            # Ambil Cupping Scores
            cupping_scores = {"ID": id_text}
            score_table = driver.find_elements(By.CLASS_NAME, "sample_information")[1]  # Tabel kedua
            rows = score_table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                headers = row.find_elements(By.TAG_NAME, "th")
                if len(cells) == 2 and len(headers) == 2:
                    cupping_scores[headers[0].text.strip()] = cells[0].text.strip()
                    cupping_scores[headers[1].text.strip()] = cells[1].text.strip()
            data_detail["Cupping Scores"].append(cupping_scores)

            # Ambil Green Analysis
            green_analysis = {"ID": id_text}
            green_table = driver.find_elements(By.CLASS_NAME, "sample_information")[2]  # Tabel ketiga
            rows = green_table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                headers = row.find_elements(By.TAG_NAME, "th")
                if len(cells) == 2 and len(headers) == 2:
                    green_analysis[headers[0].text.strip()] = cells[0].text.strip()
                    green_analysis[headers[1].text.strip()] = cells[1].text.strip()
            data_detail["Green Analysis"].append(green_analysis)

            # Ambil Certification Information
            cert_info = {"ID": id_text}
            cert_table = driver.find_elements(By.CLASS_NAME, "sample_information")[3]  # Tabel keempat
            rows = cert_table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                headers = row.find_elements(By.TAG_NAME, "th")
                if len(cells) == 1 and len(headers) == 1:
                    cert_info[headers[0].text.strip()] = cells[0].text.strip()
            data_detail["Certification Information"].append(cert_info)

            print(f"✅ Data detail ID {id_text} berhasil diambil!")
            print(sample_info)
            print(cupping_scores)
            print(green_analysis)
            print(cert_info)


        except Exception as e:
            print(f"❌ Gagal mengambil data detail ID {id_text}: {e}")

# Simpan semua data ke Excel
with pd.ExcelWriter("hasil_scraping.xlsx", engine="openpyxl") as writer:
    for sheet, df in data_main.items():
        df.to_excel(writer, sheet_name=sheet, index=False)
    for sheet, data in data_detail.items():
        pd.DataFrame(data).to_excel(writer, sheet_name=sheet, index=False)

print("🎉 Semua data berhasil disimpan ke hasil_scraping.xlsx!")

# Tutup browser
driver.quit()

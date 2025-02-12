import time
import json
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

# Gerçek Chrome'u başlat
chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
options = webdriver.ChromeOptions()
options.binary_location = chrome_path
options.add_argument("--start-maximized")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--log-level=3")
options.add_experimental_option("excludeSwitches", ["enable-logging"])

# ChromeDriver'ı başlat
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Linklerin bulunduğu metin dosyası
file_path = "link.txt"
excel_file = "fiyatlar.xlsx"

# Eğer Excel dosyası varsa oku, yoksa boş bir DataFrame oluştur
if os.path.exists(excel_file):
    df = pd.read_excel(excel_file, index_col=None)
else:
    df = pd.DataFrame(columns=["Ürün Adı", "Fiyat"])

# Ürünlerin saklanacağı sözlük
new_prices = {}

# Linkleri oku
with open(file_path, "r", encoding="utf-8") as file:
    links = file.read().splitlines()

# Her linki ziyaret et
for url in links:
    if not url.strip():
        continue

    driver.get(url)
    time.sleep(3)  # Sayfanın yüklenmesini bekle

    try:
        # Ürün adı
        product_name = driver.find_element(By.CLASS_NAME, "proName").text.strip()
        # Fiyat
        price = driver.find_element(By.CLASS_NAME, "newPrice").text.strip()

        new_prices[product_name] = price  # Yeni fiyatı kaydet

        # Eğer ürün zaten Excel'de varsa fiyatını güncelle
        if product_name in df["Ürün Adı"].values:
            old_price = df.loc[df["Ürün Adı"] == product_name, "Fiyat"].values[0]

            if old_price != price:
                print(f"🔴 Fiyat değişti! {product_name}: {old_price} → {price}")
                df.loc[df["Ürün Adı"] == product_name, "Fiyat"] = price
            else:
                print(f"✅ Fiyat değişmedi: {product_name} - {price}")

        else:
            # Ürün yoksa yeni satır olarak ekle
            print(f"🟢 Yeni ürün eklendi: {product_name} - {price}")
            df = pd.concat([df, pd.DataFrame([[product_name, price]], columns=["Ürün Adı", "Fiyat"])], ignore_index=True)

    except Exception as e:
        print(f"Hata oluştu ({url}): {e}")

# Güncellenmiş veriyi Excel'e kaydet
df.to_excel(excel_file, index=False)

# Tarayıcıyı kapat
driver.quit()

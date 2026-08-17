import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

opts = Options()
opts.add_argument("--headless")

driver = webdriver.Chrome(options=opts)
driver.get("http://10.119.105.77:3033")
time.sleep(5)

element = driver.find_elements(By.XPATH, '//table//td[1]')
list_domain = [elem.text.strip() for elem in element if elem.text.strip()]
driver.quit()

hasil = []

for domain in list_domain:
    url = domain if domain.startswith("http") else f"http://{domain}"
    try:
        res = requests.get(url, timeout=5, headers={"User-Agent": "Mozilla/5.0"})
        status = 'Bisa diakses' if res.status_code == 200 else f'HTTP {res.status_code}'
    except:
        status = 'Gagal / Down'
    hasil.append({'Domain': domain, 'Status': status})

pd.DataFrame(hasil).to_excel("hasil.xlsx", index=False)
print('Selesai! Hasil disimpan di file hasil.xlsx')
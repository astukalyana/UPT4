from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, load_workbook
import time

from user_input import get_data, get_masa

#MEMILIH WP UNTUK DIPROSES ID BILL PAT
data_wp = get_data()
print(data_wp)
nama_wp = data_wp["nama"]
npwpd = data_wp["npwpd"]
metode = data_wp["metode"]
zona = data_wp["zona"]
manfaat = data_wp["manfaat"]

#SELENIUM SIMPAD
driver = webdriver.Chrome()

driver.get("https://pajak.tangerangkab.go.id/login")

dismiss_btn = driver.find_element(By.XPATH, "//*[@id='myModal1']/div/div/div[3]/button")
dismiss_btn.click()

username = driver.find_element(By.NAME, "userid")
username.send_keys("UPTWIL4")

password = driver.find_element(By.NAME, "passwd")
password.send_keys("TOMI")
password.send_keys(Keys.RETURN)

driver.get("https://pajak.tangerangkab.go.id/pad_tangerangkab/sptpd/add/18")

form_nopd = driver.find_element(By.XPATH, '//*[@id="nopd"]')
form_nopd.send_keys(npwpd)
time.sleep(0.5)
form_nopd.send_keys(Keys.RETURN)

form_tanggal = driver.find_element(By.XPATH, '//*[@id="masadari"]')
form_tanggal.click()

datepicker_bulan = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/table/thead/tr[1]/th[2]').text
datepicker_bulan_sebelumnya = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/table/thead/tr[1]/th[1]')

masa = "Januari 2021"
while(datepicker_bulan != masa):
    datepicker_bulan_sebelumnya.click()
    datepicker_bulan = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/table/thead/tr[1]/th[2]').text

datepicker_table = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/table/tbody')
datepicker_tanggal = datepicker_table.find_elements(By.TAG_NAME, 'td')

for i in datepicker_tanggal:
    if(i.text == "1"):
        i.click()
        break

selection_carahitung = Select(driver.find_element(By.XPATH, '//*[@id="air_isflat"]'))
selection_carahitung.select_by_visible_text(metode)

selection_zona = Select(driver.find_element(By.XPATH, '//*[@id="air_zona_id"]'))
selection_zona.select_by_visible_text(zona)

selection_manfaat = Select(driver.find_element(By.XPATH, '//*[@id="air_manfaat_id"]'))
selection_manfaat.select_by_visible_text(manfaat)

form_volume = driver.find_element(By.XPATH, '//*[@id="volume"]')
form_volume.clear()
form_volume.send_keys("100")
input()
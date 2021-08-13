from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

from pandas import ExcelWriter
import pandas as pd

options = Options()
options.page_load_strategy = 'eager'

driver = webdriver.Chrome(options=options)

# Indices
indices = ['idx-composite',
           'japan-ni225',
           'shanghai-composite',
           'hang-sen-40',
           'uk-100',
           'xetra-dax-price',
           'us-30',
           'us-spx-500',
           'nq-100']

def get_index(driver,index):
    driver.find_element_by_xpath('/html/body/div[5]/section/div[8]/div[2]/select/option[2]').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[5]/section/div[9]/table[1]/tbody/tr[1]/td[7]')))
        
    data = [index]
    data.extend([driver.find_element_by_xpath('/html/body/div[5]/section/div[9]/table[1]/tbody/tr[1]/td[2]').text,
                 driver.find_element_by_xpath('/html/body/div[5]/section/div[9]/table[1]/tbody/tr[1]/td[7]').text])
    
    return(data)

index_change = []
for index in indices:
    driver.get(f"https://www.investing.com/indices/{index}-historical-data")
    index_change.append(get_index(driver, index))

index_df = pd.DataFrame(index_change, columns=['index','last','change'])

# Sectors
driver.get("https://www.idx.co.id/data-pasar/laporan-statistik/statistik/#weekly")
driver.find_element_by_xpath("/html/body/main/div[1]/div[2]/div[2]").click()
driver.find_element_by_xpath("/html/body/main/div[2]/div/div[2]/div/div[3]/table/tbody/tr[1]/td[4]/a").click()

with ExcelWriter('C:/Users/jason/Desktop/KSEP_WMR.xlsx') as writer:
    index_df.to_excel(writer, sheet_name='index', index=False)
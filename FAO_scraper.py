import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options

#### paths####
driver_path = r"C:\Users\lpicc\OneDrive\Documents\drivers\chromedriver_win32\chromedriver.exe"                 #enter driver path --> using firefox
url = "https://foodandagricultureorganization.shinyapps.io/PrevalenceofUndernourishmentProjectionTool-2020/"     #url ...founf this within page..different frame 
download_path = r"C:\Users\lpicc\OneDrive\Documents\Pardee work\viz team\FAO_scraper"
final_doc_path = r"C:\Users\lpicc\OneDrive\Documents\Pardee work\viz team\FAO_scraper\data.xlsx"

#set download folder
chromeOptions = Options()
chromeOptions.add_experimental_option("prefs", {"download.default_directory": download_path})

#set driver
driver = webdriver.Chrome(driver_path, options=chromeOptions)           
driver.get(url)

## country drop down menu select
select = driver.find_element_by_id("cou")
sel = Select(select)

#get length of country menu - how many countries are there?
countries = len(sel.options)

#download files
for i in range(1,countries):
    sel.select_by_index(i)
    time.sleep(3)
    driver.find_element_by_id("sav_bttn").click()

    
#set dataframes    
cols1 = ['Variable', 'Country', '2000', '2001', '2002', '2003', '2004', '2005', '2006',
       '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2014', '2015',
       '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024',
       '2025', '2026', '2027', '2028', '2029', '2030']
cols2 = ['Variable','Country', '1999-01', '2000-02', '2001-03', '2002-04', '2003-05',
       '2004-06', '2005-07', '2006-08', '2007-09', '2008-10', '2009-11',
       '2010-12', '2011-13', '2012-14', '2013-15', '2014-16', '2015-17',
       '2016-18', '2017-19', '2018-20', '2019-21', '2020-22', '2021-23',
       '2022-24', '2023-25', '2024-26', '2025-27', '2026-28', '2027-29',
       '2028-30', '2029-31']
df_1= pd.DataFrame(columns = cols1)
df_2 = pd.DataFrame(columns=cols2)


# loop through downloaded files
for filename in os.listdir(r"C:\Users\lpicc\OneDrive\Documents\Pardee work\viz team\FAO_scraper"):
    if filename.endswith(".csv"):
        path = r"C:\Users\lpicc\OneDrive\Documents\Pardee work\viz team\FAO_scraper\\" + filename

        #check if estimates are available
        check = pd.read_csv(path, usecols=[0])
        if "not available" in check.iloc[1,0]:
            continue
        else:
            #append data
            df = pd.read_csv(path, skiprows=9).iloc[:7,:]

            df_temp_1 = df.iloc[0:4,1:]
            df_1 = df_1.append(df_temp_1, ignore_index=True)

            df_temp_2 = df.iloc[5:,1:]
            df_temp_2.columns = cols2
            df_2 = df_2.append(df_temp_2, ignore_index = True)

#save to excel
writer = pd.ExcelWriter(final_doc_path, engine='xlsxwriter')

df_1.to_excel(writer, sheet_name= "sheet 1")
df_2.to_excel(writer, sheet_name = "sheet 2")
writer.save()
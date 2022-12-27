import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import os
import pandas as pd


def scraper_01(chrome_driver_path, data_website, year_to_select, cause_of_death):

    chrome_driver = webdriver.Chrome(chrome_driver_path)
    chrome_driver.get(data_website) # chrome_driver.execute_script("window.scrollTo(0, document.body.scrollHeight);") 

    ################################################################################################################################
    # Criteria Filtering #
    ################################################################################################################################
    print(f"Start data scraping for year={year_to_select}, cause of death={cause_of_death},  by district...")

    # Year - selectively
    year_box = chrome_driver.find_element(By.XPATH, '//*[@id="cbo0"]')
    year_option_screening = year_box.find_elements(By.TAG_NAME, "option")
    year_box_option_list = [i.get_attribute("value") for i in year_option_screening]
    year_indexing = int(year_box_option_list.index(year_to_select))
    year_box_choose_one = year_option_screening[year_indexing]
    year_box_choose_one.click(); chrome_driver.implicitly_wait(5)
    
    # Gender=all, Age=all, District=all
    gender_box = chrome_driver.find_element(By.XPATH, '//*[@id="All1"]')
    age_box = chrome_driver.find_element(By.XPATH, '//*[@id="All2"]')
    district_box = chrome_driver.find_element(By.XPATH, '//*[@id="All3"]')
    gender_box.click(); age_box.click(); district_box.click()
    chrome_driver.implicitly_wait(5)

    # Cause of death - selectively
    death_cause_box = chrome_driver.find_element(By.XPATH, '//*[@id="cbo4"]')
    death_cause_option_screening = death_cause_box.find_elements(By.TAG_NAME, "option")
    death_cause_option_list = [i.get_attribute("value") for i in death_cause_option_screening]
    death_cause_indexing = int(death_cause_option_list.index(cause_of_death))
    death_box_choose_one = death_cause_option_screening[death_cause_indexing]
    death_box_choose_one.click()

    ################################################################################################################################
    # Requesting Table #
    ################################################################################################################################
    calculation_button = chrome_driver.find_element(By.XPATH, '//*[@id="bt1"]')
    calculation_button.click()
    chrome_driver.implicitly_wait(5)

    try: 
        pop_up = Alert(chrome_driver); pop_up_text = pop_up.text
        pop_up.accept()
        data_scraped = pd.DataFrame(
            index=['Central & Western', 'Eastern (HK)', 'Southern (HK)', 'Wan Chai','Kowloon City', 'Kwun Tong', 'Sham Shui Po', 'Wong Tai Sin',
                        'Yau Tsim Mong', 'Islands', 'Kwai Tsing', 'North', 'Sai Kung','Sha Tin', 'Tai Po', 'Tsuen Wan', 
                        'Tuen Mun', 'Yuen Long', 'Marine','Outside Hong Kong', 'Unknown', 'Total'],
            columns=[f"{cause_of_death}"],
        ); data_scraped.fillna(pop_up_text, inplace=True)
        
        print(f"Done data scraping for year={year_to_select}, cause of death={cause_of_death}!")

    except:
        table_dropdown = chrome_driver.find_element(By.XPATH, '//*[@id="cboRowVar"]')
        table_dropdown_screening = table_dropdown.find_elements(By.TAG_NAME, "option")
        table_var_choose_one = table_dropdown_screening[0] ############################################## 0="Year of Death Registration"
        table_var_choose_one.click()

        show_table_button = chrome_driver.find_element(By.XPATH, '//*[@id="showhide"]/input')
        show_table_button.click()

        table = chrome_driver.find_element(By.XPATH, '//*[@id="bivContainer"]')
        index_screening = table.find_elements(By.TAG_NAME, "th")
        index_list = [i.text for i in index_screening]
        data_screening = table.find_elements(By.TAG_NAME, "td")
        data_list = [i.text for i in data_screening]

        reshaped_index = index_list[2:23]; reshaped_index.append("Total")
        reshaped_data = data_list[1:22]; reshaped_data.append(data_list[22])
        col_name = f"{cause_of_death}"

    ################################################################################################################################
    # Reshaping data #
    ################################################################################################################################
        data_scraped = pd.DataFrame(
            index=reshaped_index,
            data=reshaped_data,
            columns=[col_name])
        
        print(f"Done data scraping for year={year_to_select}, cause of death={cause_of_death}!")
    
    chrome_driver.quit()

    return data_scraped.transpose()

def scraper_01_extended (local_driver_path, hksar_page, list_of_diseases, year):
    result_list=[]
    for i in list_of_diseases:
        interim_result = scraper_01(
            chrome_driver_path=local_driver_path,
            data_website=hksar_page,
            year_to_select=year,
            cause_of_death=i)
        
        result_list.append(interim_result)
        result_for_diseases=pd.concat(result_list)
        result_for_diseases=result_for_diseases.reset_index().rename(columns={"index":year})
    
    return result_for_diseases

################################################################################################################################
# Run #
################################################################################################################################
start_time = datetime.now()
codebook = pd.read_excel("/Users/noel/Desktop/Noel/HKSAR_WebScraping/CodeBook_WebScraping_Mortality.xlsx")
disease_code_list = codebook["Code"]

output_folder = "/Users/noel/Desktop/Noel/HKSAR_WebScraping/Data"
year_enter_by_user = str(input("Please enter a year to start: "))

df = scraper_01_extended(    
    local_driver_path="/Users/noel/Desktop/Noel/HKSAR_WebScraping/chromedriver",
    hksar_page="https://www.healthyhk.gov.hk/phisweb/enquiry/mo_ysad10_indiv_e.html#resultanchor",
    year=year_enter_by_user, 
    list_of_diseases=["C33", "C34"]
); os.chdir(output_folder)

df.to_excel(f"{year_enter_by_user}_Age.xlsx", sheet_name=year_enter_by_user)

print(f"Done for {year_enter_by_user}, by district x gender!"); end_time = datetime.now()
print(f"Running time: {end_time-start_time}")
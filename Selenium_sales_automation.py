from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import logging

import win32com.client
import os
import traceback
import time
import shutil
from datetime import datetime, timedelta

print("BISMILLA-HIRRAHMAN-NIRRAHEEM")

Log_file = r"C:\AUTOMATION FILES\LOG FILE\SALES LOGGER"
Log_path = os.path.join(Log_file,f"SALES AUTOMATION {datetime.now().strftime('%d-%m-%Y--%H-%M')}.txt")
logging.basicConfig(
    filename=Log_path,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='a',
    force = True
)
logging.info("AUTOMATION STARTED")


# --------------------- DATE & TIME ---------------------
Today = datetime.now().strftime("%d-%m-%Y")
Start_Time = datetime.now()
TOD_NOS = datetime.now().strftime("%d")
Yesterday = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%y : %H:%M")
Yesternos = (datetime.now() - timedelta(days=1)).strftime("%d")

print(Yesterday, " ", Yesternos)
def SALES_DOWNLOAD(SALES_RUN, retry = 0 ):
    try:
        Sales_Data = r"C:\AUTOMATION FILES\SQL\SALES FILE"
        Sales_Return = r"C:\AUTOMATION FILES\SQL\SALES RETURN"

        try:
            def deletion_existing(DATA, NOS):
                matching_files = [
                f for f in os.listdir(DATA)
                if f.lower().startswith(NOS.lower())
                ]
                if len(matching_files) > 5:
                    for csv_file in matching_files:
                        file_path = os.path.join(DATA, csv_file)
                        os.remove(file_path)
                    print(f"Removed: {csv_file}")
                    logging.info(f"{matching_files} FOUNDED & REMOVED : {csv_file}")
                else :
                    print(len(matching_files),"ONLY Exists")
                    logging.info("%d EXISTS", len(matching_files))
        except Exception as E:
            pass

        deletion_existing(Sales_Data,NOS="521_s")
        deletion_existing(Sales_Return,NOS="559_s")  


        def wait_andget_csv(folder, before_files):

            while True:

                time.sleep(2)

                after_files = set(os.listdir(folder))
                new_files = after_files - before_files

                csv_files = [f for f in new_files
                    if f.lower().endswith(".csv")]

                if csv_files:

                    latest_file = max(
                        [os.path.join(folder, f) for f in csv_files],
                        key=os.path.getctime)

                    return latest_file
                
        def excel_fix_csv(input_csv):

            excel = win32com.client.DispatchEx("Excel.Application")

            excel.Visible = False

            try:

                excel.DisplayAlerts = False

                wb = excel.Workbooks.Open(input_csv)

                wb.SaveAs(input_csv, FileFormat=6)

                wb.Close(SaveChanges=False)

                print("EXCEL FIXED :", input_csv)

            finally:

                excel.Quit()

        try:

            prefSD = {
                'download.default_directory': Sales_Data,
                'download.prompt_download': False,
                'safebrowsing.enabled': True,
                'download.upgrade_download': True
            }
            SD_options = webdriver.ChromeOptions()

            SD_options.add_argument("--start-maximized")

            SD_options.add_argument("--disable-gpu")

            SD_options.add_argument("--disable-extensions")

            SD_options.add_argument("--disable-popup-blocking")


            SD_options.add_experimental_option('prefs',prefSD)

            Driver = webdriver.Chrome(options=SD_options)

            time.sleep(5)

            Wait_web = WebDriverWait(Driver, 40)

            Long_wait = WebDriverWait(Driver, 1100)

            GO_frugal = 'https://Companyname.gofrugal.com'

            Driver.get(GO_frugal)

            def Direction(Search_key, Search_list):

                try:

                    Wait_web.until(EC.presence_of_element_located((By.XPATH, "//input[@name='j_username']"))).send_keys("USERNAME")

                    Wait_web.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Password']"))).send_keys("PASSWORD")

                    Wait_web.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Login']"))).click()

                    time.sleep(10)

                    try:
                        Wait_web.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Hide Alert']"))).click()

                        time.sleep(5)

                    except:
                        pass

                except:
                    pass

                Wait_web.until(EC.presence_of_element_located((By.XPATH, "//input[@ng-model = 'search.selected']"))).send_keys(Search_key)

                time.sleep(10)

                Wait_web.until(EC.presence_of_element_located((By.XPATH, f"//li/a[normalize-space() = '{Search_list}']"))).click()

                time.sleep(10)

                Driver.switch_to.frame("dynamicBodyIframe")

                Long_wait.until(EC.element_to_be_clickable((By.XPATH, "//i[@class = 'icons-calendar-icomoon']"))).click()

                time.sleep(3)
                if TOD_NOS == "01":
                    Last_Month = Wait_web.until(EC.element_to_be_clickable((By.XPATH,"//button[normalize-space()='Last Month']")))
                    Last_Month.click()
                elif datetime.now().time() > datetime.strptime("16:00","%H:%M").time():    
                    Start_date = Wait_web.until(
                        EC.element_to_be_clickable(
                            (By.XPATH,"(//table[contains(@class,'table table-condensed flush')])[1]//td[normalize-space()='1']")))
                    Start_date.click()
                    time.sleep(3)
                    Till_date = Wait_web.until(
                        EC.element_to_be_clickable((By.XPATH,
                                                    f"(//table[contains(@class,'table table-condensed flush')])[2]//td[normalize-space() = {TOD_NOS}]")))
                    Till_date.click()
                    time.sleep(3)
                    Apply = Wait_web.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space() = 'Apply']")))
                    Driver.execute_script("arguments[0].click();",Apply)
                else:
                    Start_date = Wait_web.until(
                        EC.element_to_be_clickable(
                            (By.XPATH,"(//table[contains(@class,'table table-condensed flush')])[1]//td[normalize-space()='1']")))
                    Start_date.click()
                    time.sleep(3)
                    Till_date = Wait_web.until(
                        EC.element_to_be_clickable((By.XPATH,
                                                    f"(//table[contains(@class,'table table-condensed flush')])[2]//td[normalize-space() = {Yesternos}]")))
                    Till_date.click()
                    time.sleep(3)
                    Apply = Wait_web.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space() = 'Apply']")))
                    Driver.execute_script(
                        "arguments[0].click();",Apply)    
                time.sleep(5)

                Final_Apply = Wait_web.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[normalize-space()='Apply']")))

                Driver.execute_script("arguments[0].click();",Final_Apply)

                time.sleep(15)

                Long_wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//i[@class='icons-export3-icomoon']"))).click()

                time.sleep(10)

                Time_Now = datetime.now()

                Duration1 = Time_Now - Start_Time

                print(
                    "WAITED TIME FOR LOADED :",Duration1," TIME :",Time_Now)

                Long_wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH,"//div[normalize-space() = 'Export as csv']/following-sibling::button"))).click()
                    

                print("EXPORT STARTED :", Time_Now)

                time.sleep(10)
                
                Driver.switch_to.default_content()
                
            before_sales = set(os.listdir(Sales_Data))

            Direction("521","Sale - Bill Wise Item Detailed")

            sales_file = wait_andget_csv(Sales_Data,before_sales)

            excel_fix_csv(sales_file)

            print("SALES FILE READY :", sales_file)

            logging.info(f"SALES FILE READY : {sales_file}")
            Driver.switch_to.new_window('tab')
            Driver.get(GO_frugal)

            before_return = set(os.listdir(Sales_Data))

            Direction("559","Sales Return - Item Wise Detailed")

            sales_returnf = wait_andget_csv(Sales_Data,before_return)

            excel_fix_csv(sales_returnf)

            new_path = os.path.join(Sales_Return,os.path.basename(sales_returnf))

            shutil.move(sales_returnf,new_path)

            print("SALES RETURN FILE READY :", new_path)

            print("PROCESS COMPLETED SUCCESSFULLY")
            logging.info("PROCESS COMPLETED ALH--------")
        except Exception as E:

            print("ERROR ABSORBED AS :", E)

            logging.error(f"ERROR ABSORBED AS : {E}")

            logging.error(traceback.format_exc())

            traceback.print_exc()
            raise

        finally:

            try:
                Driver.quit()

            except:
                pass
    except Exception as e:
        print("ERROR OCCURED")
        if retry < 2:
            print(f"Retrying -- {retry}")
            time.sleep(5)
            SALES_DOWNLOAD(SALES_RUN, retry+1 )
        else : 
            print("ATTEMPT FAILED")

SALES_DOWNLOAD(SALES_RUN=any)  
print("THE PROCESS IS COMPLETED")
logging.info("AUTOMATION PROCESS COMPLETED -AL---------")
logging.shutdown()

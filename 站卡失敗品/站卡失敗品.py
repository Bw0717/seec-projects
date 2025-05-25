from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import pandas as pd
import keyboard
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import letter
import sys
import os
from pylibdmtx.pylibdmtx import encode
from impoSrt_os import process_files
from drive_upload import upload_to_drive
def runcard_automation():
    file_path = r'G:\MES自動化\站卡對照表.xlsx'
    current_date_prefix = "AMG" + datetime.now().strftime("%y%m%d")
    print(current_date_prefix)
    try:
        df = pd.read_excel(file_path)
        print("Excel檔案已成功讀取，欄位名稱如下：")
        print(df.columns)
        search_value = "24LMAAT45"
        matched_rows = df[df['料號'].astype(str) == search_value]
        if not matched_rows.empty:
            result_product = matched_rows['產品'].values[0]
            print(f"\n找到對應的產品：{result_product}")
        else:
            result_product = None
            print(f"找不到料號 '{search_value}' 對應的產品。")
    except FileNotFoundError:
        print(f"無法找到檔案：{file_path}")
        result_product = None
    except KeyError:
        print("資料中沒有找到 '產品' 或 '料號' 欄位，請確認 Excel 檔案格式是否正確。")
        result_product = None
    except Exception as e:
        print(f"讀取檔案時發生錯誤：{e}")
        result_product = None
    global actions2
    download_dir = os.path.abspath("./pdf_downloads")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    options = Options()
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--ignore-certificate-errors')  
    options.add_argument('--disable-web-security')      
    input_acc = "14574"
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(options=options,service=service)
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    actions2 = ActionChains(driver)
    wait = WebDriverWait(driver, 10)  
    wait2 = WebDriverWait(driver, 300)  
    try:
        driver.get("http://cimes.seec.com.tw/AMG/CimesDesktop.aspx")
        username_input = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
        original_window = driver.current_window_handle
        all_windows = driver.window_handles
        for window in all_windows:
            if window != original_window:
                driver.switch_to.window(window)
                break
        driver.close()
        driver.switch_to.window(original_window)
        password_input = driver.find_element(By.CSS_SELECTOR, "input#Password")
        login_button = driver.find_element(By.ID, "LoginButton")
        username_input.send_keys(input_acc)
        password_input.send_keys(input_acc)
        login_button.click()
        all_windows = driver.window_handles
        for window in driver.window_handles:
            if window != all_windows:
                driver.switch_to.window(window)
                break
        element_menu = wait.until(
            EC.element_to_be_clickable((By.ID, "TestMenu"))
        )
        element_menu.click()
        driver.switch_to.frame("ifmMenu")
        element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div[title="新增站卡批號"][cri="A2019022615434529060000"]')))
        element4.click()
        driver.switch_to.default_content()
        iframe = driver.find_element(By.CLASS_NAME, "win1")
        driver.switch_to.frame(iframe)
        keyboard.press_and_release('ctrl+-')
        keyboard.press_and_release('ctrl+-')
        station_number_list = []
        select_element = Select(driver.find_element("id", "ddlProduct"))
        select_element.select_by_visible_text(result_product)
        time.sleep(0.5)
        select_element2 = Select(driver.find_element("id", "ddlDevice"))
        select_element2.select_by_visible_text(search_value)
        time.sleep(0.5)
        try:
            modal_window = driver.find_element(By.XPATH, "//div[contains(@class, 'ui-dialog')]")
            error_code = modal_window.find_element(By.CSS_SELECTOR, "div.ui-dialog-titlebar.exceptionTitle")
            close_button = modal_window.find_element(By.XPATH, ".//input[@value='關閉' and @type='button']")
            close_button.click()
            driver.quit()
            os._exit(0)
        except:
            pass
        input_element = driver.find_element("id", "ttbGroupNum")
        input_element.send_keys("5")
        input_element2 = driver.find_element("id", "ttbPerNum")
        input_element2.send_keys("10")
        time.sleep(0.5)
        btnAdd = wait.until(EC.presence_of_element_located((By.ID, "btnAdd")))
        btnAdd.click()
        time.sleep(0.5)
        station_numbers = driver.find_elements(By.XPATH, "//table[@id='gvQuery']//span[starts-with(@id, 'gvQuery_ctl') and contains(@id, '_lblSubNo')]")
        for element in station_numbers:
            station_number = element.text
            station_number_list.append(station_number)
        print("站卡編號列表:", station_number_list)
        btnOK = wait.until(EC.presence_of_element_located((By.ID, "btnOK")))
        btnOK.click()
        time.sleep(0.5)
        checked_station_numbers = set()
        driver.switch_to.default_content()
        element_menu = wait.until(
            EC.element_to_be_clickable((By.ID, "TestMenu"))
        )
        element_menu.click()
        driver.switch_to.frame("ifmMenu")
        element5 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div[title="列印日工單/站卡"][cri="A2019022615452929300000"]')))
        element5.click()
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(iframes[1])
        radio_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "rbtOper"))
        )
        radio_button.click()
        time.sleep(0.5)
        input_rundard_number = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ttbDayWOOrOperation"))
        )
        input_rundard_number.clear()
        input_rundard_number.send_keys(current_date_prefix)
        time.sleep(0.5)
        btnAdd2 = wait.until(EC.presence_of_element_located((By.ID, "btnAdd")))
        btnAdd2.click()
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "tbody")))
        table_rows = driver.find_elements(By.XPATH, "//tbody/tr")
        while True:
            try:
                table_rows = driver.find_elements(By.XPATH, "//tbody/tr")
                for row in table_rows:
                    try:
                        station_number_elements = row.find_elements(By.XPATH, ".//td[2]/span")
                        if not station_number_elements:
                            continue
                        station_number = station_number_elements[0].text.strip()
                        if station_number in station_number_list and station_number not in checked_station_numbers:
                            print(f"匹配站卡編號: {station_number}")
                            try:
                                checkbox = row.find_element(By.XPATH, ".//td[1]/input[@type='checkbox']")
                                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", checkbox)
                                time.sleep(0.5) 
                                if not checkbox.is_selected():
                                    checkbox.click()
                                    print(f"已勾選站卡編號: {station_number}")
                                    checked_station_numbers.add(station_number)
                                else:
                                    print(f"站卡編號 {station_number} 已經被勾選")
                                    checked_station_numbers.add(station_number)
                            except Exception as e:
                                print(f"無法點擊 checkbox，原因: {e}")
                    except Exception as e:
                        print(f"無法點擊 checkbox，原因: {e}")
                if len(checked_station_numbers) == len(station_number_list):
                    print("所有指定站卡編號已成功勾選，結束操作。")
                    break
                driver.execute_script("window.scrollBy(0, 100);")
                time.sleep(1)
            except Exception as e:
                print(f"錯誤：{e}")
                break
        btnOK2 = wait.until(EC.presence_of_element_located((By.ID, "btnOK")))
        btnOK2.click()
        time.sleep(3)
        original_window = driver.current_window_handle
        time.sleep(2)
        all_windows = driver.window_handles
        for window in all_windows:
            if window != original_window:
                driver.switch_to.window(window)
                print("已切換到新視窗")
                keyboard.press_and_release('ctrl+s')
                time.sleep(2)
                save_path = r"G:\MES自動化\downloads\SEECRunCard.pdf"
                keyboard.write(save_path)
                time.sleep(1)
                keyboard.press_and_release('enter')
                time.sleep(1)
                if keyboard.is_pressed('alt'):  
                    print("覆蓋提示彈出，選擇覆蓋")
                    keyboard.press_and_release('alt+y')
                else:
                    keyboard.press_and_release('tab')
                    time.sleep(0.5)
                    keyboard.press_and_release('enter')
                    print("已選擇覆蓋")
                print("成功選擇路徑並確認儲存")
                time.sleep(3)

                break
    finally:
        driver.quit()





if __name__ == '__main__':
    runcard_automation()
    process_files()
    folder_id = '1DENTzbEklFWbHDhsInxIxJzjfPd0oino'  
    file_path = r'G:\MES自動化\downloads\SEECRunCard.pdf'
    share_url = upload_to_drive(file_path, folder_id)






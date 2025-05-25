import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QStackedWidget, QTabWidget, QRadioButton, QGroupBox, QCheckBox, QFileDialog
from PyQt5.QtCore import Qt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from utils.pd import read_excel_and_process
from utils.pd import read_excel_and_group_by_hierarchy
from utils.pd import read_excel_and_auto_work_order
from ping3 import ping 
import os
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter
from linebot.models import *
import tkinter as tk
from tkinter import filedialog
chrome_path = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
options = Options()
options.add_argument("--disable-usb-devices")
options.add_argument('--disable-gpu')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')  
options.add_argument('--disable-web-security')     
options.add_argument("--no-default-browser-check")
options.add_argument("--no-first-run")
options.add_argument('--headless') 
options.add_argument("--disable-features=PasswordManagerEnabled,PasswordLeakDetection,AutofillServerCommunication")
options.add_experimental_option("prefs", {
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False
})  
options.binary_location = chrome_path
input_acc = "14575"
service = Service(ChromeDriverManager().install())

def lot(DEPARTMENT,df,workstation_name):
    try:
        driver = webdriver.Chrome(options=options,service=service)
        driver.set_window_size(1920, 1080)
        driver.maximize_window()
        wait = WebDriverWait(driver, 10)  
        wait2 = WebDriverWait(driver, 2) 
        driver.get(f"http://cimes.seec.com.tw/{DEPARTMENT}/CimesDesktop.aspx")
        time.sleep(0.5)
        main_window = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main_window:
                new_window = handle
                break
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(main_window)
        username_input = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
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
        op_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(op_element).click().perform()
        element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="WipRule_Batch"]')))
        element4.click()
        button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="批次變更工作站"]')))
        actions.move_to_element(button).click().perform()
        driver.switch_to.default_content()
        iframe = driver.find_element(By.CLASS_NAME, "win1")
        driver.switch_to.frame(iframe)
        for idx, row in df.iterrows():
            batch_number = row["批號"]
            input_element3 = wait.until(EC.element_to_be_clickable((By.ID, "ttbLot")))
            input_element3.clear()
            input_element3.send_keys(batch_number)
            input_element3.send_keys(Keys.RETURN)
            time.sleep(0.5)
            try:
                error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                time.sleep(0.5)
                error_text = error_element.text
                df.at[idx, "執行結果"] = error_text 
                close_button = wait2.until(EC.element_to_be_clickable(
                    (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                ))
                driver.execute_script("arguments[0].click();", close_button)
                continue
            except:
                df.at[idx, "執行結果"] = "執行成功" 
      
        select_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='csReason']//select[@id='_ddlcsReason']"))
        )
        select2 = Select(select_element)   
        select2.select_by_index(1)
        time.sleep(0.5)
        table = driver.find_element(By.XPATH, "//span[text()='批號']/ancestor::table")
        select_element2 = table.find_element(By.ID, "ddlNewOperation")
        select_obj = Select(select_element2)
        found = False
        for option in select_obj.options:
            if workstation_name in option.text:
                print(f"✅ 選中：{option.text}")
                select_obj.select_by_visible_text(option.text)
                found = True
                break
        if not found:
            print(f"找不到包含{workstation_name}的選項")
        OK_btn = table.find_element(By.ID, "btnOK")
        OK_btn.click()
    finally:
        time.sleep(2)
        driver.quit      


def lot_2(DEPARTMENT,df,workstation_name):
    try:
        driver = webdriver.Chrome(options=options,service=service)
        driver.set_window_size(1920, 1080)
        driver.maximize_window()
        wait = WebDriverWait(driver, 10)  
        wait2 = WebDriverWait(driver, 2) 
        driver.get(f"http://cimes.seec.com.tw/{DEPARTMENT}/CimesDesktop.aspx")
        time.sleep(0.5)
        main_window = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main_window:
                new_window = handle
                break
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(main_window)
        username_input = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
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
        op_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(op_element).click().perform()
        element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="WipRule_Batch"]')))
        element4.click()
        button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="批次變更流程"]')))
        actions.move_to_element(button).click().perform()
        driver.switch_to.default_content()
        iframe = driver.find_element(By.CLASS_NAME, "win1")
        driver.switch_to.frame(iframe)
        for idx, row in df.iterrows():
            batch_number = row["批號"]
            input_element3 = wait.until(EC.element_to_be_clickable((By.ID, "ttbLot")))
            input_element3.clear()
            input_element3.send_keys(batch_number)
            input_element3.send_keys(Keys.RETURN)
            time.sleep(0.5)
            try:
                error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                time.sleep(0.5)
                error_text = error_element.text
                df.at[idx, "執行結果"] = error_text 
                close_button = wait2.until(EC.element_to_be_clickable(
                    (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                ))
                driver.execute_script("arguments[0].click();", close_button)
                continue
            except:
                df.at[idx, "執行結果"] = "執行成功" 
        select_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='csReason']//select[@id='_ddlcsReason']"))
        )
        select2 = Select(select_element)   
        select2.select_by_index(1)
        time.sleep(0.5)
        table = driver.find_element(By.XPATH, "//span[text()='批號']/ancestor::table")
        select_element2 = table.find_element(By.ID, "ddlNewRoute")
        select_obj = Select(select_element2)
        select_obj.select_by_visible_text(workstation_name)
        time.sleep(0.5)
        table = driver.find_element(By.XPATH, "//span[text()='批號']/ancestor::table")
        select_element3 = table.find_element(By.ID, "ddlRouteVersion")
        select_obj2 = Select(select_element3)
        keyword = "線上版本"
        found = False
        for option in select_obj2.options:
            if keyword in option.text:
                print(f"選中：{option.text}")
                select_obj2.select_by_visible_text(option.text)
                found = True
                break
        if not found:
            print("找不到包含『線上版本』的選項")
        time.sleep(0.5)
        OK_btn = table.find_element(By.ID, "btnOK")
        OK_btn.click()
    finally:
        time.sleep(2)
        driver.quit
              


def lot_3(DEPARTMENT,df,workstation_name):
    try:
        driver = webdriver.Chrome(options=options,service=service)
        driver.set_window_size(1920, 1080)
        driver.maximize_window()
        wait = WebDriverWait(driver, 10)  
        wait2 = WebDriverWait(driver, 2) 
        driver.get(f"http://cimes.seec.com.tw/{DEPARTMENT}/CimesDesktop.aspx")
        time.sleep(0.5)
        main_window = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main_window:
                new_window = handle
                break
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(main_window)
        username_input = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
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
        op_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(op_element).click().perform()
        element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="WipRule_Batch"]')))
        element4.click()
        button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="批次變更屬性"]')))
        actions.move_to_element(button).click().perform()
        driver.switch_to.default_content()
        iframe = driver.find_element(By.CLASS_NAME, "win1")
        driver.switch_to.frame(iframe)
        radio_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "rbtSYS"))
        )
        if not radio_btn.is_selected():
            radio_btn.click()
        for idx, row in df.iterrows():
            batch_number = row["批號"]
            input_element3 = wait.until(EC.element_to_be_clickable((By.ID, "ttbLot")))
            input_element3.clear()
            input_element3.send_keys(batch_number)
            input_element3.send_keys(Keys.RETURN)
            time.sleep(0.5)
            try:
                error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                time.sleep(0.5)
                error_text = error_element.text
                df.at[idx, "執行結果"] = error_text 
                close_button = wait2.until(EC.element_to_be_clickable(
                    (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                ))
                driver.execute_script("arguments[0].click();", close_button)
                continue
            except:
                df.at[idx, "執行結果"] = "執行成功" 
        select_element = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[@id='csReason']//select[@id='_ddlcsReason']"))
        )
        select2 = Select(select_element)   
        select2.select_by_index(1)
        time.sleep(0.5)
        table = driver.find_element(By.XPATH, "//span[text()='批號']/ancestor::table")
        select_element2 = table.find_element(By.ID, "ddlAttribute")
        select_obj = Select(select_element2)
        select_obj.select_by_visible_text("WO")
        time.sleep(0.5)
        table = driver.find_element(By.XPATH, "//span[text()='批號']/ancestor::table")
        input_element = table.find_element(By.ID, "ttbNewValue")
        input_element.send_keys(str(workstation_name))
        time.sleep(0.5)
        OK_btn = table.find_element(By.ID, "btnOK")
        OK_btn.click()
    finally:
        time.sleep(2)
        driver.quit



def lot_4(DEPARTMENT,df,new_product,workstation_name):
    try:
        driver = webdriver.Chrome(options=options,service=service)
        driver.set_window_size(1920, 1080)
        driver.maximize_window()
        wait = WebDriverWait(driver, 10)  
        wait2 = WebDriverWait(driver, 2) 
        driver.get(f"http://cimes.seec.com.tw/{DEPARTMENT}/CimesDesktop.aspx")
        time.sleep(0.5)
        main_window = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main_window:
                new_window = handle
                break
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(main_window)
        username_input = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
        password_input = driver.find_element(By.CSS_SELECTOR, "input#Password")
        login_button = driver.find_element(By.ID, "LoginButton")
        username_input.send_keys(input_acc)
        password_input.send_keys(input_acc)
        login_button.click()
        first_run = True
        for idx, row in df.iterrows():        
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
            op_element = wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
            )
            actions = ActionChains(driver)
            actions.move_to_element(op_element).click().perform()
            element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="WipRule_Modify"]')))
            element4.click()
            button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="變更型號"]')))
            actions.move_to_element(button).click().perform()
            driver.switch_to.default_content()
            if first_run:
                driver.switch_to.frame(0)
                first_run = False
            else:
                driver.switch_to.frame(1)
            time.sleep(0.5)
            batch_number = row["批號"]
            input_element3 = wait.until(EC.element_to_be_clickable((By.ID, "CimesInputBox")))
            input_element3.clear()
            try:
                error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                time.sleep(0.5)

                close_button = wait2.until(EC.element_to_be_clickable(
                    (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                ))
                driver.execute_script("arguments[0].click();", close_button)
            except:
                pass
            input_element3 = wait.until(EC.element_to_be_clickable((By.ID, "CimesInputBox")))
            input_element3.send_keys(batch_number)
            time.sleep(0.5)
            input_element3.send_keys(Keys.RETURN)
            time.sleep(0.5)
            try:
                error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                time.sleep(0.5)
                error_text = error_element.text
                df.at[idx, "執行結果"] = error_text 
                close_button = wait2.until(EC.element_to_be_clickable(
                    (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                ))
                driver.execute_script("arguments[0].click();", close_button)
                continue
            except:
                time.sleep(0.5)
                input_element4 = wait.until(EC.element_to_be_clickable((By.ID, "ttbDeviceFilter")))
                input_element4.clear()
                try:
                    error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                    time.sleep(0.5)

                    close_button = wait2.until(EC.element_to_be_clickable(
                        (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                    ))
                    driver.execute_script("arguments[0].click();", close_button)
                except:
                    pass
                select_element = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//div[@id='csProduct']//select[@id='_ddlcsProduct']"))
                )
                select10 = Select(select_element)   
                select10.select_by_visible_text(new_product)
                time.sleep(0.5)
                input_element4 = wait.until(EC.element_to_be_clickable((By.ID, "ttbDeviceFilter")))
                time.sleep(0.5)
                input_element4.send_keys(str(workstation_name))
                input_element4.send_keys(Keys.RETURN)
                try:
                    error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                    time.sleep(0.5)
                    error_text = error_element.text
                    df.at[idx, "執行結果"] = error_text 
                    close_button = wait2.until(EC.element_to_be_clickable(
                        (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                    ))
                    driver.execute_script("arguments[0].click();", close_button)
                    continue
                except:
                    pass

                select_element = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//div[@id='csReason']//select[@id='_ddlcsReason']"))
                )
                select2 = Select(select_element)   
                select2.select_by_index(1)
                time.sleep(0.5)
                OK_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@id='UpdatePanel1']//a[@id='btnOK']"))
                )

                OK_btn.click()
                df.at[idx, "執行結果"] = "執行成功" 
    finally:
        time.sleep(2)
        driver.quit


def auto_work_order(file_path,dele_type): 
    data_inputs = read_excel_and_auto_work_order(file_path)
    print(data_inputs)
    try:
        driver = webdriver.Chrome(options=options,service=service)
        driver.set_window_size(1920, 1080)
        driver.maximize_window()
        wait = WebDriverWait(driver, 10)   
        driver.get(f"http://cimes.seec.com.tw/{dele_type}/Security/CimesUserLogin.aspx")
        time.sleep(0.5)
        main_window = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main_window:
                new_window = handle
                break
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(main_window)

        username_input = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
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

        # 
        element_menu.click()
        driver.switch_to.frame("ifmMenu")
        op_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(op_element).click().perform()

        element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="CustRule"]')))
        element4.click()

        button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="日工單開立"]')))
        actions.move_to_element(button).click().perform()
        
        time.sleep(3)
        driver.switch_to.default_content()
        iframe = driver.find_element(By.CLASS_NAME, "win1")
        driver.switch_to.frame(iframe)
        
        for data_input in data_inputs:
            input_element3 = wait.until(EC.presence_of_element_located((By.ID, "CimesInputBox")))
            input_element3.clear()
            time.sleep(1)
            input_element3.send_keys(data_input)
            input_element3.send_keys(Keys.RETURN)
            time.sleep(1.5)
            enable_input = wait.until(EC.presence_of_element_located((By.ID, "ttbUnCreateQty")))
            value = enable_input.get_attribute('value')
            if value == '0':
                continue
            else:
                pass
            input_element4 = wait.until(EC.presence_of_element_located((By.ID, "ttbDayWOQty")))
            input_element4.send_keys(value)
            add_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "btnAdd"))
            )                
            add_button.click()

            ok_button = wait.until(
            EC.element_to_be_clickable((By.ID, "btnOK"))
            )             
            actions.move_to_element(ok_button).click().perform()
            time.sleep(1)    
    finally:
        print("執行完成")
        driver.quit()

def automation_input(file_path,dele_type,check):
    grouped_data = read_excel_and_group_by_hierarchy(file_path)
    print(grouped_data)
    error_messages = []
    try:
        driver = webdriver.Chrome(options=options,service=service)
        driver.set_window_size(1920, 1080)
        driver.maximize_window()
        actions2 = ActionChains(driver)
        wait = WebDriverWait(driver, 10)  
        driver.get(f"http://cimes.seec.com.tw/{dele_type}/Security/CimesUserLogin.aspx")
        time.sleep(0.5)
        main_window = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main_window:
                new_window = handle
                break
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(main_window)
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
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
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "TestMenu"))
        )

        # 
        element.click()
        driver.switch_to.frame("ifmMenu")
        element3 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#navFRE.ctgr"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(element3).click().perform()
        element4 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPFRE"][cimesclass="EDC"]')))
        element4.click()

        element5 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#caption')))
        element5.click()

        driver.switch_to.default_content()
        iframe = driver.find_element(By.CLASS_NAME, "win1")
        driver.switch_to.frame(iframe)


        dropdown = driver.find_element(By.ID, "ddlProduct")
        select = Select(dropdown)
        select.select_by_value("ALL")
        time.sleep(1)
        for product, stations in grouped_data.items():
            input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ttbDeviceFilter"))
            )
            input_field.clear()
            input_field.send_keys(product)


            element6 = driver.find_element(By.CSS_SELECTOR, '#btnRefresh[name="btnRefresh"]')
            actions2 = ActionChains(driver)

            actions2.move_to_element(element6).click().perform()
            time.sleep(1)
            for station, parameters in stations.items():

                dropdown = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "ddlOperation"))
                )

                select = Select(dropdown)
                select.select_by_value(station)
                time.sleep(1)

                add_button = driver.find_element(By.ID, "btnAdd")
                actions2 = ActionChains(driver)
                actions2.move_to_element(add_button).click().perform()

            
                time.sleep(1)
                second_iframe = driver.find_element(By.XPATH, '//iframe[contains(@src, "EDCOperSetPara.aspx")]')  # 用 xpath 定位第二層 iframe
                driver.switch_to.frame(second_iframe)  # 切換到第二層 iframe

                try:
                    ddl_rule_name_element = driver.find_element(By.ID, "ddlRuleName")
                    select = Select(ddl_rule_name_element)
                    select.select_by_index(1)
                except:
                    pass
                time.sleep(1)
                ddl_type_element = driver.find_element(By.ID, "ddlType")
                select = Select(ddl_type_element)
                select.select_by_index(1)
                time.sleep(1.5)
                dropdown_element = Select(driver.find_element(By.ID, "ddlCorelationOper"))
                dropdown_element.select_by_visible_text(station)

                #刪除動作
                if check == 1:  
                    while True:
                        elements = driver.find_elements(By.CSS_SELECTOR, ".CSGridEditButton")
                        if elements:
                            elements[0].click()
                            enable_time_input = WebDriverWait(driver, 20).until(
                                EC.element_to_be_clickable((By.XPATH, "//div[@id='PanelParameterInfo']//input[@id='ttbEnableTime']"))
                            )                        
                            ActionChains(driver).move_to_element(enable_time_input).perform()
                            value = enable_time_input.get_attribute('value')



                            ttbDisableTime = wait.until(EC.element_to_be_clickable((By.ID, "ttbDisableTime")))
                            ttbDisableTime.clear()
                            time.sleep(0.5)
                            ttbDisableTime.send_keys(value)
                            ttbDisableTime.send_keys(Keys.RETURN)
                            time.sleep(0.5)
                            ok_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.ID, "btnOKParameter"))
                            )                
                            ok_button.click()
                            time.sleep(1)
                            save_button5 = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.ID, "btnSave"))
                            )                
                            save_button5.click()
                            time.sleep(1)
                        else:
                            break  
                else:
                    pass


                for parameter in parameters:
                    time.sleep(1)
                    driver.execute_script("document.querySelector('#btnAdd').click();")
                    button11 = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnQueryParameter"))
                    )
                    driver.execute_script("arguments[0].click();", button11)

                    input_element = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "ttbFilter"))
                    )
                    input_element.clear()
                    time.sleep(1)
                    input_element2 = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "ttbFilter"))
                    )
                    input_element2.click
                    input_element2.send_keys(parameter["參數名稱"])
                    time.sleep(0.5)
                    input_element2.send_keys(Keys.RETURN)
                    time.sleep(2)
                    try:
                        element_by_css = WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "a[href=\"javascript:__doPostBack('gvFilterItems','Select$0')\"]"))
                        )
                        driver.execute_script("arguments[0].click();", element_by_css)
                        time.sleep(0.5)
                        input_element = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "ttbOperSeq"))
                        )
                        input_element.clear()
                        input_element.send_keys(parameter["順序"])

                        dropdown3 = driver.find_element(By.ID, "ddlOperCritical")
                        select3 = Select(dropdown3)
                        select3.select_by_value(parameter["重要性"])

                        datetime_input2 = driver.find_element(By.ID, "ttbEnableTime")
                        datetime_input2.clear()
                        datetime_input2.send_keys("2020/12/25 17:42:12")
                        datetime_input2.send_keys(Keys.RETURN)
                        time.sleep(1)
                        btnOKParameter = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "btnOKParameter"))
                        )
                        btnOKParameter.click()
                    except:
                        error_messages.append("找不到" + parameter["參數名稱"])
                        btnClose = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "btnClose"))
                        )
                        btnClose.click()
                        time.sleep(1)
                        continue
                    try:
                        error_message = WebDriverWait(driver, 1).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "span#lblParaMsg.CSMust"))
                        )
                        error_text = error_message.text
                        error_messages.append(error_message.text)

                        time.sleep(1)
                        btnCCParameter = WebDriverWait(driver, 2).until(
                            EC.element_to_be_clickable((By.ID, "btnCloseParameter"))
                        )
                        btnCCParameter.click()
                        time.sleep(1)
                    except:
                        save_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "btnSave"))
                        )
                        save_button.click() 
                        time.sleep(0.5)



                time.sleep(2)

                exit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "btnExit"))
                )
                exit_button.click()

                alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert.accept()

                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)
                time.sleep(1)
        with open("error_messages.txt", "w", encoding="utf-8") as file:
            for message in error_messages:
                file.write(message + "\n")
        time.sleep(5)
    finally:
        print("執行完成")
        driver.quit()


def automation_parameter(file_path,check,dele_type):
    data = read_excel_and_process(file_path)
    print(data)
    edit_data = []
    error_messages = []
    try:
        driver = webdriver.Chrome(options=options,service=service)
        driver.set_window_size(1920, 1080)
        driver.maximize_window()
        actions2 = ActionChains(driver)
        wait = WebDriverWait(driver, 10)  
        wait2 = WebDriverWait(driver, 2) 
        driver.get(f"http://cimes.seec.com.tw/{dele_type}/CimesDesktop.aspx")
        time.sleep(0.5)
        main_window = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main_window:
                new_window = handle
                break
        driver.switch_to.window(new_window)
        driver.close()
        driver.switch_to.window(main_window)
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
        )
        password_input = driver.find_element(By.CSS_SELECTOR, "input#Password")
        login_button = driver.find_element(By.ID, "LoginButton")

        # 
        username_input.send_keys(input_acc)
        password_input.send_keys(input_acc)
        login_button.click()
        all_windows = driver.window_handles

        for window in driver.window_handles:
            if window != all_windows:
                driver.switch_to.window(window)
                break
        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "TestMenu"))
        )
        element.click()
        driver.switch_to.frame("ifmMenu")
        element3 = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#navFRE.ctgr"))
        )
        actions = ActionChains(driver)
        actions.move_to_element(element3).click().perform()

        element4 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPFRE"][cimesclass="EDC"]')))
        element4.click()


        button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="工程資料參數維護"]')))
        ActionChains(driver).move_to_element(button).click().perform()


        driver.switch_to.default_content()
        iframe = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "win1")))
        driver.switch_to.frame(iframe)

        add_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnAdd")))
        actions2 = ActionChains(driver)
        actions2.move_to_element(add_button).click().perform()


        for item in data:
            time.sleep(1)
            dropdown_element2 = wait.until(EC.element_to_be_clickable((By.ID, "ddlDataType")))
            select_value2 = item["data_type"]
            script = f"""
                arguments[0].value = '{select_value2}';
                arguments[0].dispatchEvent(new Event('change'));
            """
            driver.execute_script(script, dropdown_element2)
            time.sleep(0.5)
            input_element6 = wait.until(EC.presence_of_element_located((By.ID, "ttbSamplesize")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element6, item["get_quan"])

            input_element1 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#ttbParameter')))
            driver.execute_script("arguments[0].value = arguments[1];", input_element1, item["param_name"])
            
            input_element2 = wait.until(EC.presence_of_element_located((By.ID, "ttbDisplayName")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element2, item["param_name2"])
            
            input_element3 = wait.until(EC.presence_of_element_located((By.ID, "ttbTarget")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element3, item["fit"])
            
            input_element4 = wait.until(EC.presence_of_element_located((By.ID, "ttbUSL")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element4, item["limit_high"])
            
            input_element5 = wait.until(EC.presence_of_element_located((By.ID, "ttbLSL")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element5, item["limit_low"])
            


            rbt_enable = driver.find_element(By.ID, "rbtEnable")
            rbt_disable = driver.find_element(By.ID, "rbtDisable")

            if  item["state"] == "啟用":
                rbt_enable.click()  
            else:
                rbt_disable.click()  


            dropdown_element = wait.until(EC.element_to_be_clickable((By.ID, "ddlUnit")))
            select_value = item["unit"]
            script = f"""
                arguments[0].value = '{select_value}';
                arguments[0].dispatchEvent(new Event('change'));
            """
            driver.execute_script(script, dropdown_element)
            time.sleep(0.5)

            


            if select_value2 == "Number":
                dropdown_element3 = wait.until(EC.element_to_be_clickable((By.ID, "ddlVariableType")))
                select = Select(dropdown_element3)
                select.select_by_index("1")

            time.sleep(1)

                
            save_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnSave")))
            actions3 = ActionChains(driver)
            actions3.move_to_element(save_button).click().perform()

            time.sleep(3)
            
            try:
                error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
                error_text = error_element.text
                error_messages.append(error_text)

                error_info = {
                        "param_name":item["param_name"],
                        "data_type": item["data_type"],
                        "param_name2": item["param_name2"],
                        "fit": item["fit"],
                        "limit_high": item["limit_high"],
                        "limit_low": item["limit_low"],
                        "state": item["state"],
                        "unit": item["unit"],
                        "get_quan": item["get_quan"]
                }
                print(error_info)
                edit_data.append(error_info)
                print(edit_data)
                
                close_button = wait2.until(EC.element_to_be_clickable(
                    (By.XPATH, "//td//input[@value='關閉' and @type='button']")
                ))
                driver.execute_script("arguments[0].click();", close_button)
                continue
            except:
                pass

            add_button2 = wait.until(EC.element_to_be_clickable((By.NAME, "btnAdd")))
            actions5 = ActionChains(driver)
            actions5.move_to_element(add_button2).click().perform()

        exit_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnExit")))
        actions4 = ActionChains(driver)
        actions4.move_to_element(exit_button).click().perform()
        time.sleep(1)

        print(edit_data)





        if check == 1:

            for item in edit_data:
                ttbParameter = wait.until(EC.element_to_be_clickable((By.NAME, "ttbParameter")))
                ttbParameter.clear()
                actions.move_to_element(ttbParameter).click().perform()
                ttbParameter.send_keys(item["param_name"])
                btnQuery = wait.until(EC.element_to_be_clickable((By.NAME, "btnQuery")))
                btnQuery.click()
                time.sleep(2)
                EditButton = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @class='CSGridEditButton']"))
                )
                actions.move_to_element(EditButton).click().perform()
                

                time.sleep(1)
                dropdown_element2 = wait.until(EC.element_to_be_clickable((By.ID, "ddlDataType")))
                select_value2 = item["data_type"]
                script = f"""
                    arguments[0].value = '{select_value2}';
                    arguments[0].dispatchEvent(new Event('change'));
                """
                driver.execute_script(script, dropdown_element2)
                time.sleep(0.5)
                input_element6 = wait.until(EC.presence_of_element_located((By.ID, "ttbSamplesize")))
                driver.execute_script("arguments[0].value = arguments[1];", input_element6, item["get_quan"])


                
                input_element2 = wait.until(EC.presence_of_element_located((By.ID, "ttbDisplayName")))
                driver.execute_script("arguments[0].value = arguments[1];", input_element2, item["param_name2"])
                
                input_element3 = wait.until(EC.presence_of_element_located((By.ID, "ttbTarget")))
                driver.execute_script("arguments[0].value = arguments[1];", input_element3, item["fit"])
                
                input_element4 = wait.until(EC.presence_of_element_located((By.ID, "ttbUSL")))
                driver.execute_script("arguments[0].value = arguments[1];", input_element4, item["limit_high"])
                
                input_element5 = wait.until(EC.presence_of_element_located((By.ID, "ttbLSL")))
                driver.execute_script("arguments[0].value = arguments[1];", input_element5, item["limit_low"])
                


                rbt_enable = driver.find_element(By.ID, "rbtEnable")
                rbt_disable = driver.find_element(By.ID, "rbtDisable")

                if  item["state"] == "啟用":
                    rbt_enable.click()  
                else:
                    rbt_disable.click()  


                dropdown_element = wait.until(EC.element_to_be_clickable((By.ID, "ddlUnit")))
                select_value = item["unit"]
                script = f"""
                    arguments[0].value = '{select_value}';
                    arguments[0].dispatchEvent(new Event('change'));
                """
                driver.execute_script(script, dropdown_element)
                time.sleep(0.5)

                


                if select_value2 == "Number":
                    dropdown_element3 = wait.until(EC.element_to_be_clickable((By.ID, "ddlVariableType")))
                    select = Select(dropdown_element3)
                    select.select_by_index("1")

                time.sleep(1)

                
                save_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnSave")))
                actions3 = ActionChains(driver)
                actions3.move_to_element(save_button).click().perform()

                time.sleep(2)
                exit_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnExit")))
                actions4 = ActionChains(driver)
                actions4.move_to_element(exit_button).click().perform()
                time.sleep(2)
        else:
            pass

        time.sleep(5)
    finally:
        with open("error_messages.txt", "w", encoding="utf-8") as file:
            for message in error_messages:
                file.write(message + "\n")
        driver.quit()



class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("MES批量自動化")
        self.resize(600, 400)
        global check
        check = 0
        self.stacked_widget = QStackedWidget(self)

        self.main_page = self.create_main_page()
        self.param_page = self.create_param_page()

        self.stacked_widget.addWidget(self.main_page)
        self.stacked_widget.addWidget(self.param_page)

        layout = QVBoxLayout()
        layout.addWidget(self.stacked_widget)
        self.setLayout(layout)

        footer_label = QLabel("Version: 1.0.0 | Writer: CHIH-HSIANG", self)
        footer_label.setAlignment(Qt.AlignRight)  
        layout.addWidget(footer_label, alignment=Qt.AlignBottom | Qt.AlignRight) 

    def create_main_page(self):
        """主頁面：包含兩個按鈕"""
        main_page = QWidget()
        

        label = QLabel("歡迎使用批量自動化工具", main_page)
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 20px; font-weight: bold; color: #333;")


        button1 = QPushButton("參數資料", main_page)
        button1.setFixedSize(200, 100)
        button1.setStyleSheet("font-size: 16px; background-color: #4CAF50; color: white; border-radius: 10px;")
        button1.clicked.connect(self.on_button1_click)


        button2 = QPushButton("結束", main_page)
        button2.setFixedSize(200, 100)
        button2.setStyleSheet("font-size: 16px; background-color: #FF5722; color: white; border-radius: 10px;")
        button2.clicked.connect(self.on_button2_click)


        layout = QVBoxLayout()
        layout.addWidget(label)
        

        button_layout = QHBoxLayout()
        button_layout.addWidget(button1)
        button_layout.addWidget(button2)
        button_layout.setSpacing(20)
        layout.addLayout(button_layout)


        main_page.setLayout(layout)

        return main_page
    def on_create_work_order(self):
        """處理建立工單按鈕點擊事件"""

        print("開始建立工單...")
    def create_param_page(self):
        """參數資料頁面：點擊按鈕進入的界面"""
        param_page = QWidget()


        tab_widget = QTabWidget(param_page)


        tab1 = QWidget()
        tab1_layout = QVBoxLayout()
        label1 = QLabel("請選擇一個選項:", tab1)
        group_box = QGroupBox("選擇單位", tab1)
        group_box2 = QGroupBox("選擇要執行的服務", tab1)

        group_box2.resize(180, 130)
        group_box2.move(300, 110)

        self.radio_arg = QRadioButton("ARG", group_box)
        self.radio_amg = QRadioButton("AMG", group_box)
        self.radio_am2 = QRadioButton("AM2", group_box)

        self.radio_P = QRadioButton("建立參數", group_box2)
        self.radio_I = QRadioButton("匯入參數", group_box2)


        self.radio_arg.setChecked(False)
        self.radio_P.setChecked(False)
        self.radio_arg.toggled.connect(self.on_radio_button_changed)
        self.radio_amg.toggled.connect(self.on_radio_button_changed)
        self.radio_am2.toggled.connect(self.on_radio_button_changed)
        self.radio_P.toggled.connect(self.on_radio_button_changed2)
        self.radio_I.toggled.connect(self.on_radio_button_changed2)


        radio_layout = QHBoxLayout()
        radio_layout.addWidget(self.radio_arg)
        radio_layout.addWidget(self.radio_amg)
        radio_layout.addWidget(self.radio_am2)
        group_box.setLayout(radio_layout)

        radio_layout2 = QVBoxLayout()
        radio_layout2.addWidget(self.radio_P)
        radio_layout2.addWidget(self.radio_I)
        group_box2.setLayout(radio_layout2)

        self.selected_label = QLabel("目前選擇: ", tab1)
        self.checkbox = QCheckBox("覆蓋已建立參數", tab1)
        self.checkbox.setChecked(False)
        self.checkbox.toggled.connect(self.on_checkbox_toggled)


        file_button = QPushButton("選擇檔案路徑", tab1)
        file_button.setFixedSize(200, 50)
        file_button.setStyleSheet("font-size: 14px; background-color: #FF5722; color: white; border-radius: 10px;")
        file_button.clicked.connect(self.on_select_file)


        tab1_layout.addWidget(label1)
        tab1_layout.addWidget(group_box)
        tab1_layout.addWidget(self.selected_label)
        tab1_layout.addWidget(self.checkbox)
        tab1_layout.addWidget(file_button)
        tab1.setLayout(tab1_layout)


        tab2 = QWidget()
        tab2_layout = QVBoxLayout()
        label2 = QLabel("建立工單", tab2)


        group_box2 = QGroupBox("選擇單位", tab2)
        self.radio_arg2 = QRadioButton("ARG", group_box2)
        self.radio_amg2 = QRadioButton("AMG", group_box2)
        self.radio_am22 = QRadioButton("AM2", group_box2)


        self.radio_arg2.setChecked(False)


        radio_layout2 = QHBoxLayout()
        radio_layout2.addWidget(self.radio_arg2)
        radio_layout2.addWidget(self.radio_amg2)
        radio_layout2.addWidget(self.radio_am22)
        group_box2.setLayout(radio_layout2)

        self.selected_label2 = QLabel("目前選擇:", tab2)
        self.radio_arg2.toggled.connect(self.on_radio_button_changed3)
        self.radio_amg2.toggled.connect(self.on_radio_button_changed3)
        self.radio_am22.toggled.connect(self.on_radio_button_changed3)


        file_button2 = QPushButton("選擇檔案路徑", tab2)
        file_button2.setFixedSize(200, 50)
        file_button2.setStyleSheet("font-size: 14px; background-color: #FF5722; color: white; border-radius: 10px;")
        file_button2.clicked.connect(self.on_select_file2)


        tab2_layout.addWidget(label2)
        tab2_layout.addWidget(group_box2)
        tab2_layout.addWidget(file_button2)
        tab2.setLayout(tab2_layout)





        tab3 = QWidget()
        tab3_layout = QVBoxLayout()
        label3 = QLabel("批次變更設定", tab3)

        group_box3 = QGroupBox("選擇單位", tab3)
        self.radio_arg3 = QRadioButton("ARG", group_box3)
        self.radio_amg3 = QRadioButton("AMG", group_box3)
        self.radio_am23 = QRadioButton("AM2", group_box3)

        group_box4 = QGroupBox("選擇批次功能", tab3)
        self.radio_lot1 = QRadioButton("批次工作站", group_box3)
        self.radio_lot2 = QRadioButton("批次變更工單", group_box3)
        self.radio_lot3 = QRadioButton("批次變更件號", group_box3)
        self.radio_lot4 = QRadioButton("批次變更流程", group_box3)
        self.radio_lot1.setChecked(False)

        radio_layout4 = QHBoxLayout()
        radio_layout4.addWidget(self.radio_lot1)
        radio_layout4.addWidget(self.radio_lot2)
        radio_layout4.addWidget(self.radio_lot3)
        radio_layout4.addWidget(self.radio_lot4)
        group_box4.setLayout(radio_layout4)

        self.radio_lot1.toggled.connect(self.on_radio_button_changedlot)
        self.radio_lot2.toggled.connect(self.on_radio_button_changedlot)
        self.radio_lot3.toggled.connect(self.on_radio_button_changedlot)
        self.radio_lot4.toggled.connect(self.on_radio_button_changedlot)

        radio_layout3 = QHBoxLayout()
        radio_layout3.addWidget(self.radio_arg3)
        radio_layout3.addWidget(self.radio_amg3)
        radio_layout3.addWidget(self.radio_am23)
        group_box3.setLayout(radio_layout3)


        self.selected_label3 = QLabel("目前選擇:", tab3)
        self.radio_arg3.toggled.connect(self.on_radio_button_changed4)
        self.radio_amg3.toggled.connect(self.on_radio_button_changed4)
        self.radio_am23.toggled.connect(self.on_radio_button_changed4)


        file_button3 = QPushButton("選擇檔案路徑", tab3)
        file_button3.setFixedSize(200, 50)
        file_button3.setStyleSheet("font-size: 14px; background-color: #FF5722; color: white; border-radius: 10px;")
        file_button3.clicked.connect(self.on_select_file3)

        tab3_layout.addWidget(label3)
        tab3_layout.addWidget(group_box3)
        tab3_layout.addWidget(group_box4)
        tab3_layout.addWidget(file_button3)
        tab3.setLayout(tab3_layout)






        tab_widget.addTab(tab1, "建立參數")
        tab_widget.addTab(tab2, "建立工單")  
        tab_widget.addTab(tab3, "批次變更")

        button_layout = QHBoxLayout()

        start_button = QPushButton("開始執行", param_page)
        start_button.setFixedSize(180, 70)
        start_button.setStyleSheet("font-size: 16px; background-color: #4CAF50; color: white; border-radius: 10px;")
        start_button.clicked.connect(self.on_start_button_click2)

        back_button = QPushButton("返回主頁", param_page)
        back_button.setFixedSize(180, 70)
        back_button.setStyleSheet("font-size: 16px; background-color: #f44336; color: white; border-radius: 10px;")
        back_button.clicked.connect(self.on_back_button_click)

        button_layout.addWidget(start_button)
        button_layout.addWidget(back_button)
        button_layout.setSpacing(20)

        layout = QVBoxLayout()
        layout.addWidget(tab_widget)
        layout.addLayout(button_layout)
        param_page.setLayout(layout)

        return param_page



    def on_button1_click(self):
        """切換到參數資料頁面"""
        self.stacked_widget.setCurrentIndex(1)

    def on_button2_click(self):
        """開立工單功能"""
        QApplication.quit()
    def on_radio_button_changed2(self):
        global select_type2
        select_type2 = None
        if self.radio_P.isChecked():
            select_type2 = "建立參數"
        elif self.radio_I.isChecked():
            select_type2 = "匯入參數"

    def on_radio_button_changedlot(self):
        global select_type2
        select_type2 = None
        if self.radio_lot1.isChecked():
            select_type2 = "批次工作站"
        elif self.radio_lot2.isChecked():
            select_type2 = "批次變更工單"
        elif self.radio_lot3.isChecked():
            select_type2 = "批次變更件號"
        elif self.radio_lot4.isChecked():
            select_type2 = "批次變更流程"       

    def on_radio_button_changed(self):
        global dele_type
        dele_type = None
        """處理圓形選項按鈕的選擇改變"""
        if self.radio_arg.isChecked():
            self.selected_label.setText("目前選擇: ARG")
            dele_type = "ARG"
        elif self.radio_amg.isChecked():
            self.selected_label.setText("目前選擇: AMG")
            dele_type = "AMG"
        elif self.radio_am2.isChecked():
            self.selected_label.setText("目前選擇: AM2")
            dele_type = "AM2"


    def on_checkbox_toggled(self):
        global check
        check = None
        """勾選框狀態改變"""
        if self.checkbox.isChecked():
            print("已選擇覆蓋已建立參數")
            check = 1
        else:
            print("未選擇覆蓋已建立參數")
            check = 0
    def on_select_file(self):
        """選擇檔案功能"""
        global file_path
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(self, "選擇檔案", "", "All Files (*.*)")
        if file_path:
            print(f"選擇的檔案路徑是: {file_path}")
        
    def on_select_file2(self):
        """選擇檔案功能"""
        global file_path2
        global select_type2
        file_dialog = QFileDialog(self)
        file_path2, _ = file_dialog.getOpenFileName(self, "選擇檔案", "", "All Files (*.*)")
        if file_path2:
            print(f"選擇的檔案路徑是: {file_path2}")
            select_type2 = "開立工單"

    def on_select_file3(self):
        global file_path3
        file_dialog = QFileDialog(self)
        file_path3, _ = file_dialog.getOpenFileName(self, "選擇檔案", "", "All Files (*.*)")
        if file_path3:
            print(f"選擇的檔案路徑是: {file_path3}")

    def on_start_button_click2(self):
        """開始執行按鈕"""
        print("開始執行")
        if select_type2 == "批次工作站":
            df = pd.read_excel(file_path3, sheet_name="批次工作站")
            df["執行結果"] = df["執行結果"].astype("str")
            workstation_name = df["工作站"].dropna().iloc[0]
            print(df)
            new_path = os.path.splitext(file_path3)[0] + "_執行結果.xlsx"
            lot(DEPARTMENT,df,workstation_name)
            df.to_excel(new_path, sheet_name="批次工作站", index=False)
            print(f"執行完成，並另存為：{new_path}")
        elif select_type2 == "批次變更工單":
            df = pd.read_excel(file_path3, sheet_name="批次變更工單")
            df["執行結果"] = df["執行結果"].astype("str")
            workstation_name = df["新工單"].dropna().iloc[0]
            print(df)
            new_path = os.path.splitext(file_path3)[0] + "_執行結果.xlsx"
            lot_3(DEPARTMENT,df,workstation_name)
            df.to_excel(new_path, sheet_name="批次變更工單", index=False)
            print(f"執行完成，並另存為：{new_path}")   
        elif select_type2 == "批次變更件號":
            df = pd.read_excel(file_path3, sheet_name="批次變更件號")
            df["執行結果"] = df["執行結果"].astype("str")
            new_product = df["新產品"].dropna().iloc[0]
            workstation_name = df["新件號"].dropna().iloc[0]
            print(df)
            new_path = os.path.splitext(file_path3)[0] + "_執行結果.xlsx"
            lot_4(DEPARTMENT,df,new_product,workstation_name)
            df.to_excel(new_path, sheet_name="批次變更件號", index=False)
            print(f"執行完成，並另存為：{new_path}")
        elif select_type2 == "批次變更流程":
            df = pd.read_excel(file_path3, sheet_name="批次變更流程")
            df["執行結果"] = df["執行結果"].astype("str")
            workstation_name = df["新流程"].dropna().iloc[0]
            print(df)
            new_path = os.path.splitext(file_path3)[0] + "_執行結果.xlsx"
            lot_2(DEPARTMENT,df,workstation_name)
            df.to_excel(new_path, sheet_name="批次變更流程", index=False)
            print(f"執行完成，並另存為：{new_path}")

        elif select_type2 == "建立參數":
            try:
                automation_parameter(file_path = file_path,check = check,dele_type = dele_type)
            except:
                print("系統異常")
        elif select_type2 == "匯入參數":
            try:
                automation_input(file_path = file_path,dele_type = dele_type,check = check)
            except:
                print("系統異常")
        elif select_type2 == "開立工單":
            try:
                print(file_path2)
                auto_work_order(file_path=file_path2,dele_type=dele_type)
                print("執行成功！")
            except:
                print("開立工單，執行失敗")



    def on_back_button_click(self):
        self.stacked_widget.setCurrentIndex(0)

    def on_radio_button_changed3(self):
        global dele_type
        if self.radio_arg2.isChecked():
            self.selected_label2.setText("目前選擇: ARG")
            dele_type = "ARG"
        elif self.radio_amg2.isChecked():
            self.selected_label2.setText("目前選擇: AMG")
            dele_type = "AMG"
        elif self.radio_am22.isChecked():
            self.selected_label2.setText("目前選擇: AM2")
            dele_type = "AM2"

    def on_radio_button_changed4(self):
        global DEPARTMENT
        if self.radio_arg3.isChecked():
            self.selected_label3.setText("目前選擇: ARG")
            DEPARTMENT = "ARG"
        elif self.radio_amg3.isChecked():
            self.selected_label3.setText("目前選擇: AMG")
            DEPARTMENT = "AMG"
        elif self.radio_am23.isChecked():
            self.selected_label3.setText("目前選擇: AM2")
            DEPARTMENT = "AM2"
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
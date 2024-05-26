# coding=UTF-8

import time
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import atexit

accountIndex = 0

username = ''
password = '.'

# 代理服务器
proxyHost = "115.204.167.13"
proxyPort = "50059"
proxyUser = "moVfGjvC"
proxyPass = "bN6XTLoe"
proxyType = 'https'  # socks5
driver = None

# chromedriver  path
driver_path = Service('C:\Program Files\Google\Chrome\Application\chromedriver.exe')

# 文件名称
file_path = './绍兴.xlsx'
# 获取当前时间并格式化为字符串
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
codeListPath = [
    "/html/body/div/div/div[3]/div/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div/span",
    "/html/body/div/div/div[3]/div[1]/div[3]/div/div[3]/div[2]/div[2]/div[1]/div/div[2]/table/tbody/tr[4]/td[2]/div/span",
    "/html/body/div/div/div[3]/div/div[1]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[1]/div/span"
]

detail_list = [
    "/html/body/div/div[2]/div/div[2]/section/main/div[2]/div[2]/div/div/div[3]/div[2]/div[1]/div[1]/a",
    "/html/body/div/div[2]/div/div[2]/section/main/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div[1]/a",
    "/html/body/div/div[2]/div/div[2]/section/main/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/div[1]/a",
    "/html/body/div/div[2]/div/div[2]/section/main/div[3]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div[1]/a"
]

button_list = [
    "/html/body/div/div[1]/div/div[2]/div/div/button",
    "/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div"
]


def exec_excel():
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    modified_sheets = {}
    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(f"./{file_path}.xlsx", sheet_name=sheet_name)
            try:
                count = 0
                for i in df.index:
                    print(f"公司名称,count:{count}", df.iloc[i, 1])
                    if df.iloc[i, 1] is not None and df.iloc[i, 1] != '' and not pd.isnull(df.iloc[i, 13]):
                        continue
                    df.iloc[i, 13] = get_code(str(df.iloc[i, 1]))
                    count += 1
            except Exception as e:
                print("exec_excel Exception:", e)
            modified_sheets[sheet_name] = df
        except Exception as e:
            print("read_excel Exception:", e)

    # 创建一个新的ExcelWriter对象
    with pd.ExcelWriter(f"./{sheet_name}_{current_time}.xlsx", engine='openpyxl') as writer:
        # 将修改后的DataFrame写回到Excel
        for sheet_name, df in modified_sheets.items():
            print(sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def get_data(driver, keyword):
    code = ""
    main = ""
    i = 0
    j = 0
    try:
        driver.get("https://www.tianyancha.com/advance/search/e-pc_searchinfo")
        # 首页点击查询
        time.sleep(5)

        element = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div[1]/form/div/input")
        element.clear()
        element.send_keys(keyword)
        time.sleep(5)
        search_button = retry(driver, button_list)
        if search_button is None:
            return ''
        # search_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div")
        search_button.click()
        time.sleep(5)
        detail = retry(driver, detail_list)
        if detail is None:
            return ''
        detail.click()

        time.sleep(5)
        for handle in driver.window_handles:
            if i == 0:
                main = handle
            i += 1
            if i == 2:
                driver.switch_to.window(handle)
                i = 0
                time.sleep(5)
                while j < 2:
                    try:
                        code = driver.find_element(By.XPATH, codeListPath[i]).text
                        if code is not None and code != "":
                            break
                    except NoSuchElementException:
                        print("social code index not find count:", j)
                        j = j + 1
                driver.close()
                driver.switch_to.window(main)
                print(f"{keyword}:{code}")
                break
    except Exception as e:
        for handle in driver.window_handles:
            if i == 0:
                main = handle
                driver.switch_to.window(handle)
        time.sleep(20)
        print("get_data", e)

    return code


def retry(driver, list):
    n = 0
    element = None
    # 详情页
    while n < len(list):
        try:
            element = driver.find_element(By.XPATH, list[n])
            break
        except NoSuchElementException:
            print("index not find count:", n)
            n = n + 1
    return element


def get_code(keyword):
    global driver
    global accountIndex
    if driver is None:
        option = webdriver.ChromeOptions()
        option.add_experimental_option('excludeSwitches', ['enable-automation'])  # webdriver防检测
        option.add_argument("--disable-blink-features=AutomationControlled")
        option.add_argument("--no-sandbox")
        # option.add_argument("--headless")
        option.add_argument("--disable-dev-usage")
        option.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 1})

        # 配置代理
        option.add_argument("--proxy-type=%s" % proxyType, )
        option.add_argument(f"--proxy-server={proxyHost}:{proxyPort}")
        driver = webdriver.Chrome(service=driver_path, options=option)
        driver.get("https://www.tianyancha.com/")
        time.sleep(10)

    return get_data(driver, keyword)


def exit_handler():
    global driver
    if driver is not None:
        driver.quit()
        print("Browser closed.")


# 注册退出处理程序
atexit.register(exit_handler)

if __name__ == '__main__':
    exec_excel()
    if driver is not None:
        driver.quit()

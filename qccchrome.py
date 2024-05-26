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
password = ''

# 代理服务器
proxyHost = "183.128.180.74"
proxyPort = "50088"
proxyUser = "moVfGjvC"
proxyPass = "bN6XTLoe"
proxyType = 'https'  # socks5
driver = None

# chromedriver  path
driver_path = Service('C:\Program Files\Google\Chrome\Application\chromedriver.exe')

# 文件名称
file_path = './嘉兴.xlsx'
# 获取当前时间并格式化为字符串
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
codeListPath = [
    "/html/body/div/div[2]/div[2]/div[3]/div/div[2]/div/table/tr[1]/td[3]/div/div[3]/div[1]/span[3]/span/span/span[1]",
    "/html/body/div/div[2]/div[2]/div[4]/div/div[2]/div/table/tr[1]/td[3]/div/div[3]/div[1]/span[4]/span/span/span[1]",
    "/html/body/div/div[2]/div[2]/div[3]/div/div[2]/div/table/tr[1]/td[3]/div/div[3]/div[1]/span[4]/span/span/span[1]",
    "/html/body/div/div[2]/div[2]/div[3]/div/div[2]/div/table/tr[2]/td[3]/div/div[3]/div[1]/span[4]/span/span/span[1]"
]


def exec_excel():
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    modified_sheets = {}
    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(f"./{sheet_name}.xlsx", sheet_name=sheet_name)
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
            break
        except Exception as e:
            print("read_excel Exception:", e)

    # 创建一个新的ExcelWriter对象
    with pd.ExcelWriter(f"./{sheet_name}_{current_time}.xlsx", engine='openpyxl') as writer:
        # 将修改后的DataFrame写回到Excel
        for sheet_name, df in modified_sheets.items():
            print(sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def login(driver):
    print("begin login")
    driver.delete_all_cookies()
    check(driver)
    url = "https://www.qcc.com/weblogin?back=%2F"
    driver.get(url)
    time.sleep(5)

    # 点击密码登入
    driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div[3]/img').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div[1]/div[1]/div[2]/a').click()
    time.sleep(2)

    # 输入账号密码
    driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div[1]/div[3]/form/div[1]/input').send_keys(
        username)
    time.sleep(2)
    driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div[1]/div[3]/form/div[2]/input').send_keys(
        password)

    check(driver)
    driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div[1]/div[3]/form/div[4]/button')

    # 点击登录
    time.sleep(2)

    check(driver)
    element = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div[1]/div[3]/form/div[4]/button')
    # driver.execute_script("arguments[0].click()", element)
    check(driver)
    element.click()
    time.sleep(20)


def get_data(driver, keyword):
    print("begin get social code")
    # driver.get("https://www.qcc.com")
    time.sleep(5)
    try:
        # 首页
        check(driver)
        driver.find_element(By.XPATH, '/html/body/div/div[2]/section[1]/div/div/div/div[1]/div/div/input').send_keys(
            keyword)
        time.sleep(2)
        check(driver)
        driver.find_element(By.XPATH, '/html/body/div/div[2]/section[1]/div/div/div/div[1]/div/div/span/button').click()
    except (NoSuchElementException,):
        try:
            # 搜索页
            check(driver)
            search_input = driver.find_element(By.XPATH, "/html/body/div/div[1]/div/div[1]/div/div/div/div/input")
            search_input.clear()
            time.sleep(2)
            search_input.send_keys(keyword)
            time.sleep(2)
            check(driver)
            driver.find_element(By.XPATH, "/html/body/div/div[1]/div/div[1]/div/div/div/div/span/button").click()
        except Exception:
            raise NoSuchElementException
    i = 0
    while i < 3:
        try:
            res = driver.find_element(By.XPATH, codeListPath[i]).text
            if res is not None and res != "":
                break
        except NoSuchElementException:
            print("social code index not find count:", i)
            i = i + 1

    if i >= 2 and res is None and res == "":
        raise NoSuchElementException
    try:
        print("social code :", res)
    except UnboundLocalError:
        res = ''
    time.sleep(2)
    return res


def check(driver):
    try:
        element = driver.find_element(By.XPATH, '/html/body/div/div/div/a')
        driver.execute_script("arguments[0].click();", element)
        element.click()
        time.sleep(3)
        print("check")
    except NoSuchElementException:
        print("check not found")


def get_code(keyword):
    global driver
    global username
    global password
    global accountIndex
    if driver is None:
        option = webdriver.ChromeOptions()
        option.add_experimental_option('excludeSwitches', ['enable-automation'])  # webdriver防检测
        option.add_argument("--disable-blink-features=AutomationControlled")
        option.add_argument("--no-sandbox")
        option.add_argument("--disable-dev-usage")
        # option.add_argument("--headless")
        option.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 1})
        option.add_argument("--proxy-type=%s" % proxyType, )
        option.add_argument(f"--proxy-server={proxyHost}:{proxyPort}")

        driver = webdriver.Chrome(service=driver_path, options=option)
        driver.get('http://httpbin.org/ip')  # 访问一个IP回显网站，查看代理配置是否生效了
        print("page source", driver.page_source)
    try:
        driver.find_element(By.XPATH, '/html/body/div/div[1]/div/div[1]/nav[2]/ul/li[9]/div[1]/a/img')
    except NoSuchElementException as e:
        print("is not login:", e)
        time.sleep(1)
        check(driver)
        login(driver)
    try:
        return get_data(driver, keyword)
    except NoSuchElementException as e:
        print("NoSuchElementException change account :", e)
        print("账号需验证码登录，切换账号 当前account:", account[accountIndex])
        return ''
        accountIndex = accountIndex + 1
        if accountIndex >= len(account):
            accountIndex = 0
        username = account[accountIndex][0]
        password = account[accountIndex][1]
        get_code(keyword)
    except Exception as e:
        print("Exception occurred:", e)
        return ''


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

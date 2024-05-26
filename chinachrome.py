# -- coding: utf-8
import json
import requests
import time
from datetime import datetime
import pandas as pd
from selenium.webdriver.chrome.service import Service

proxyHost = ""
proxyPort = ""
proxyType = 'https'  # socks5

# chromedriver path
driver_path = Service('C:\Program Files\Google\Chrome\Application\chromedriver.exe')

# read excel path
file_path = './台州_2024-05-26_00-00-44.xlsx'

# 信用中国 请求接口
url = 'https://public.creditchina.gov.cn/private-api/catalogSearch'

# 信用中国接口请求参数
payload = {
    "keyword": "",
    "searchState": "2",
    "tableName": "credit_xyzx_tyshxydm",
    "scenes": "defaultscenario",
    "entityType": "1,2,4,5,6,7,8",
    "page": "1"
}
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def exec_excel():
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    print(f"proxyHost:{proxyHost}"
          f"file_path:{file_path}")
    modified_sheets = {}
    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            try:
                count = 0
                for i in df.index:
                    if df.iloc[i, 1] is not None and df.iloc[i, 1] != '' and not pd.isnull(df.iloc[i, 13]):
                        continue
                    df.iloc[i, 13] = get_code(str(df.iloc[i, 1]))
                    print(f'{count}--{df.iloc[i, 1]}:{df.iloc[i, 13]}')
                    count += 1
                    time.sleep(5)
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


def get_code(keyword):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36'
        }
        # proxyMeta = f"http://{proxyHost}:{proxyPort}"
        # proxies = {
        #     "http": proxyMeta,
        #     # "https": proxyMeta
        # }
        # print(requests.get('http://httpbin.org/ip', proxies=proxies).text)
        # response_data = requests.get(url, payload, headers=headers, proxies=proxies)

        payload.update({"keyword": keyword})
        # 配置代理
        response_data = requests.get(url, payload, headers=headers)
        # 将JSON字符串解析为字典
        data = json.loads(response_data.text)
        if data and 'data' in data and 'list' in data['data'] and len(data['data']['list']) > 0:
            first_record = data['data']['list'][0]
            if 'tyshxydm' in first_record:
                first_record_tyshxydm = first_record['tyshxydm']
                return first_record_tyshxydm
            else:
                # print("第一条记录中没有 tyshxydm 字段")
                return ''
        else:
            # print("没有找到任何记录")
            return ''
    except Exception as e:
        # print("Exception get_code", e)
        print("data form", response_data.text)


if __name__ == '__main__':
    exec_excel()

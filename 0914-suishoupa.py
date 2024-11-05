import time
from selenium import webdriver
from selenium.webdriver.remote.file_detector import LocalFileDetector
import xlrd
import os
import requests

# 读取素材
def xls_duqu_suishou(info_xls_sucai):
    xls_read = xlrd.open_workbook(r"D:\xxx.xls")
    xls_sheet = xls_read.sheet_by_name("xxx-明细")
    nrows = xls_sheet.nrows
    ncols = xls_sheet.ncols
    # time.sleep(3)  等设置待时间

    # 依次读取指定文件内容
    for r in range(1,nrows):
        middle = {}
        for c,key in zip(range(ncols),["xuhao","danhao","title","photo","address","content"]):
            x = xls_sheet.cell(r,c).value
            middle[key] = x
        info_xls_sucai.append(middle)
    return info_xls_sucai


# 下载logo模块
def download(info_xls_sucai):
    root = "D://随后拍素材总//"
    for key, val in enumerate(info_xls_sucai):
        photo_dir = val.get("photo")
        photo_dir_jpg = photo_dir.split("/")
        path = root + photo_dir_jpg[-1]
        try:
            if not os.path.exists(root):
                os.mkdir(root)
            if not os.path.exists(path):
                r = requests.get(photo_dir)
                with open(path, 'wb') as f:
                    f.write(r.content)
                    f.close()
                    print(path + '文件保存成功')
            else:
                print('文件已存在')
        except:
            print('爬取失败')


def main(info_xls_sucai):
    url = "https://www.12345.suzhou.com.cn/weixin/#/"
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.maximize_window()
    driver.get(url)
    for key, val in enumerate(info_xls_sucai):
        danhao_data = []
        # 现场照片？
        photo_dir = val.get("photo")
        photo_dir_jpg = photo_dir.split("/")
        path = photo_dir_jpg[-1]
        driver.file_detector = LocalFileDetector()
        driver.find_element_by_css_selector(
            "#SearchForm > div > div:nth-child(1) > div > div > div > div > input").send_keys("D://xxx//" + path)

        # 地址 address
        address_real = val.get("address")
        address = driver.find_element_by_css_selector("#address")
        time.sleep(1)
        address.clear()
        time.sleep(1)
        address.send_keys(address_real)

        # 标题 title
        title_real = val.get("title")
        driver.find_element_by_css_selector("#title").send_keys(title_real)
        time.sleep(1)

        # 问题 question
        question = val.get("content")
        driver.find_element_by_css_selector("#question").send_keys(question)
        time.sleep(1)

        # 地区 area
        driver.find_element_by_css_selector("#SearchForm > div > div:nth-child(5) > div > div:nth-child(8)").click()
        time.sleep(1)
        driver.find_element_by_css_selector("#phone").send_keys("xxxxxxx")
        # driver.find_element_by_css_selector("#phone").send_keys("1xxxxx")
        driver.find_element_by_css_selector("#hqyzm").click()
        print("shuru yanzhema")
        time.sleep(1)

        #单号打印
        print(val.get("xuhao") +": " +  val.get("danhao") + " 已完成录入！")
        danhao_data.append(val.get("danhao"))

        print("准备下一张工单")
        driver.refresh()
        time.sleep(2)

if __name__ == '__main__':
    info_xls_sucai = []
    xls_duqu_suishou(info_xls_sucai)
    main(info_xls_sucai)
    # download(info_xls_sucai)

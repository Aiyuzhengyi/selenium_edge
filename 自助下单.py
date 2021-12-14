
# PART-TWO
# 读取素材--下载图片

import xlrd
import xlwt
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os

# 保存网站cookies，过期自动刷新。
def get_system_cookies():
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get('https://d.weidian.com/weidian-pc/weidian-loader/#/pc-vue-fx-fx-item-manage/list')
    phone = driver.find_element(By.CSS_SELECTOR,
        '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div.user-telephone > div > div > div > input')
    phone.send_keys('18662214242')
    key = driver.find_element(By.CSS_SELECTOR,
        '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div:nth-child(2) > div > div > input')
    key.send_keys('qwql0528')
    submit = driver.find_element(By.CSS_SELECTOR,
        '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div:nth-child(4) > div > button')
    submit.click()
    time.sleep(3)
    # Get all available cookies
    cookies_List = driver.get_cookies()
    xls = xlwt.Workbook()
    sheet = xls.add_sheet("weidian_cookies_update")
    sheet.write(0, 0, "cookies")
    sheet.write(1, 0, str(cookies_List))
    xls.save("d:/weidian-quwei/微店登录cookies.xls")
    driver.quit()

# 1、打开浏览器，利用本地的cookies文件登录进去，等待下一步操作。
def cookies_login():
    # Adds the cookie into current browser context
    driver_login.get('https://d.weidian.com/weidian-pc/weidian-loader/#/pc-vue-fx-fx-item-manage/list')
    xls_read = xlrd.open_workbook_xls("d:/weidian-quwei/微店登录cookies.xls")
    xls_sheet = xls_read.sheet_by_name("weidian_cookies_update")
    cookies_List = xls_sheet.cell(1, 0).value
    cookies_json1 = eval(cookies_List)
    for i in cookies_json1:
        driver_login.add_cookie(i)
    driver_login.get('https://d.weidian.com/weidian-pc/weidian-loader/#/pc-vue-fx-fx-item-manage/list')
    # print(cookies_json1)
    time.sleep(3)
    return driver_login

# 2、读取商品文件--返回含有字典的列表。
# 一次读取，多次利用
def xls_duqu():
    xls_read = xlrd.open_workbook(r"D:\weidian-quwei\微店商品库存详细.xls")
    xls_sheet = xls_read.sheet_by_name("当日库存情况")
    nrows = xls_sheet.nrows
    ncols = xls_sheet.ncols
    # 依次读取指定文件内容
    for r in range(1, nrows):
        middle = {}
        for c, key in zip(range(ncols), ["xuhao", "lianjie", "biaoti", 'jiage', 'lirun', 'kucun', 'ziyuan']):
            x = xls_sheet.cell(r, c).value
            middle[key] = x
        info_duqu_xls.append(middle)
    return info_duqu_xls

# 3、通过输入的商品id，循序查找列表中匹配的值。如果匹配到，返回对应的字段.
# 网页循环打开下单模块.
def parselweb(lianjie):
    driver_login.get(lianjie)
    # print(cookies_json1)
    time.sleep(3)
    update_lianjie_list = list(lianjie)
    update_lianjie_list.insert(8,'shop1766652488.v.')
    update_lianjie_list_news = ''.join(update_lianjie_list)
    if driver_login.current_url == update_lianjie_list_news:
        print("恭喜维斯布鲁克-利用cookies 登录成功！")
        driver_login.find_element()
    else:
        get_system_cookies()
        cookies_login()
        print("注意：维斯布鲁克-已更新cookies 再次尝试cookies登录成功！")
        driver_login.find_element(By.XPATH,'//*[@id="foot-container"]/div/div/span[2]/span[2]').click()
    return driver_login



# open url
def xiadan():
    xls_duqu()

    while True:
        xiaoshou_id = input("请输入商品ID：")
        for key, val in enumerate(info_duqu_xls):
            lianjie = val.get("lianjie")
            biaoti = val.get("biaoti")
            jiage = val.get('jiage')
            lirun = val.get('lirun')
            kucun = val.get('kucun')
            if biaoti == xiaoshou_id:
                print('查询列表得到的商品信息： 商品链接'+lianjie, biaoti,jiage ,'利润：' + lirun ,'  库存：' + kucun)
                parselweb(lianjie)  #调用第一步打开的浏览器，输入字段中含有的商品链接，等待下一步操作。
                time.sleep(10)


if __name__ == '__main__':
    info_duqu_xls = []
    root = "D:/weidian-quwei/"
    if not os.path.exists(root):
        os.mkdir(root)
    else:
        print('提示：文件夹已存在')
    driver_login = webdriver.Edge(executable_path='msedgedriver.exe')
    cookies_login()
    xiadan()


# PART-TWO
# 读取素材--下载图片

import xlrd
import xlwt
from selenium import webdriver
import time

def get_system_cookies():
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get('https://d.weidian.com/weidian-pc/weidian-loader/#/pc-vue-fx-fx-item-manage/list')
    phone = driver.find_element_by_css_selector(
        '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div.user-telephone > div > div > div > input')
    phone.send_keys('18662214242')
    key = driver.find_element_by_css_selector(
        '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div:nth-child(2) > div > div > input')
    key.send_keys('qwql0528')
    submit = driver.find_element_by_css_selector(
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

def parselweb(lianjie):
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get(lianjie)
    def cookies_login():
        # Adds the cookie into current browser context
        xls_read = xlrd.open_workbook_xls("d:/weidian-quwei/微店登录cookies.xls")
        xls_sheet = xls_read.sheet_by_name("weidian_cookies_update")
        cookies_List = xls_sheet.cell(1, 0).value
        cookies_json1 = eval(cookies_List)
        for i in cookies_json1:
            driver.add_cookie(i)
        driver.get(lianjie)
        print(cookies_json1)
        time.sleep(3)
    cookies_login()
    update_lianjie_list = list(lianjie)
    update_lianjie_list.insert(8,'shop1766652488.v.')
    update_lianjie_list_news = ''.join(update_lianjie_list)
    if driver.current_url == update_lianjie_list_news:
        print("恭喜维斯布鲁克-利用cookies 登录成功！")
    else:
        get_system_cookies()
        cookies_login()
        print("注意：维斯布鲁克-已更新cookies 再次尝试cookies登录成功！")



def xls_duqu():
    xls_read = xlrd.open_workbook(r"D:\weidian-quwei\微店商品库存详细.xls")
    xls_sheet = xls_read.sheet_by_name("当日库存情况")
    nrows = xls_sheet.nrows
    ncols = xls_sheet.ncols
    # time.sleep(3)  等设置待时间

    # 依次读取指定文件内容
    for r in range(1, nrows):
        middle = {}
        for c, key in zip(range(ncols), ["xuhao", "lianjie", "biaoti", 'jiage', 'lirun', 'kucun', 'ziyuan']):
            x = xls_sheet.cell(r, c).value
            middle[key] = x
        info_duqu_xls.append(middle)
    return info_duqu_xls

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
                print(lianjie,biaoti,jiage ,'利润：' + lirun ,'  库存：' + kucun)
                parselweb(lianjie)


if __name__ == '__main__':
    info_duqu_xls = []
    xiadan()

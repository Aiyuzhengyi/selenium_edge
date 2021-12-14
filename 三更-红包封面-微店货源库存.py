import os
import re
import time
import requests
import xlrd
import xlwt
from selenium import webdriver
from selenium.webdriver.common.by import By
'''三更主要特点：增加webdriver如何获取Cookie，已经利用Cookie登陆。
这里使用到了官方文档：https://www.selenium.dev/zh-cn/documentation/webdriver/browser/cookies/中提到的添加Cookie和获取Cookies。
driver.add_cookie({"name": "test1", "value": "cookie1"}) 和 print(driver.get_cookies())。
Python中一切皆为对象，既然生成了一个driver，也可以作为参数传给其他函数调用哦。

日期：2021年12月10日16:27:28
作者:爱与正义
用途：微店当天库存获取，导出excel。
'''


# 1.【selenium模块】遥控win10自带的Edge浏览器打开指定网页（输入账号密码，根据逻辑进行操作）。
# 2.特点：无论网站有多少页，获取网页中总计多少条，利用正则表达式提取int型数字并取整，即翻页次数。处理如下：（num = re.findall('\d+', text)和 nub_new = int(num[0]) // 100。
# 3.特点：无论单个网页有多少数量，通过获取指定范围所有的tr标签。 for循环每个tr标签内，利用xpath索取tr标签内相应字段。（这里注意，存在一种情况：selenium提取不了标签文本。如果提取的元素文本为空，这是可能就是定位的
#     元素被隐藏了，即需要判断是否被隐藏。怎么解决？通过textContent, innerText, innerHTML等属性获取。我选择的是innerHTML。）函数完成后retur 列表，后续待用。
# 4.【xlwt模块】写入info数据,利用time模块依据当天日期存储指定路径。
# 5.为方便获取网站上所有的图片，还写了两个函数：读取素材 和 下载图片 。（注意的是，由于path中有'/'，需要清除以防保存报错。

def get_system_cookies():
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get(url)
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


def parselweb():
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get(url)
    def cookies_login():
        # Adds the cookie into current browser context
        xls_read = xlrd.open_workbook_xls("d:/weidian-quwei/微店登录cookies.xls")
        xls_sheet = xls_read.sheet_by_name("weidian_cookies_update")
        cookies_List = xls_sheet.cell(1, 0).value
        cookies_json1 = eval(cookies_List)
        for i in cookies_json1:
            driver.add_cookie(i)
        driver.get(url)
        print(cookies_json1)
        time.sleep(3)
    cookies_login()

    if driver.current_url == url:
        print("恭喜维斯布鲁克-利用cookies 登录成功！")
    else:
        get_system_cookies()
        cookies_login()
        print("注意：维斯布鲁克-已更新cookies 再次尝试cookies登录成功！")
    # 以上为驱动浏览器打开相应的网址，输入对应账号密码登陆后获取cookies保存到本地excel，实测发现无论商家后台还是买家端都可以共享cookies。如果cookies失效，调用该段函数更新cookies。

    # 单页设定100条每页+翻页次数设定。
    driver.find_element(By.XPATH,'//*[@id="weidianHelp"]/div/div[3]/div[2]/div[2]/span[2]/div/div/input').click()
    time.sleep(3)
    driver.find_element(By.XPATH,'/html/body/div[4]/div[1]/div[1]/ul/li[5]').click()
    time.sleep(4)
    text = driver.find_element(By.CSS_SELECTOR,
        '#weidianHelp > div > div.card > div.fx-seller-item-table > div.el-pagination > span.el-pagination__total').text
    text_num = re.findall('\d+', text)
    text_num_int = int(text_num[0]) // 100
    text_num_int_fanye = text_num_int + 1

    # 实现循环
    for i in range(text_num_int_fanye):
        # 局部变量单次循环后清空操作
        href_text = []
        img_text = []
        name_text = []
        price_text = []
        profit_text = []
        stock_text = []
        trs = driver.find_elements(By.XPATH,
            '//*[@id="weidianHelp"]/div/div[3]/div[2]/div[1]/div[4]/div[2]/table/tbody/tr')
        for tr in trs:
            # 商品链接
            href = tr.find_element_by_xpath('td[2]/div/div/a').get_attribute('href')
            href_text.append(href)

            # 商品图片
            img = tr.find_element_by_xpath('td[2]/div/div/a/img').get_attribute('src')
            img_text.append(img)

            # 商品标题
            name = tr.find_element_by_xpath('td[2]/div/div/div[2]/p[1]').get_attribute('innerHTML')
            name_text.append(name)

            # 供销商定价
            price = tr.find_element_by_xpath('td[3]/div/p[1]').get_attribute('innerHTML')
            price_text.append(price)

            # 利润
            profit = tr.find_element_by_xpath('td[4]/div/span').get_attribute('innerHTML')
            profit_text.append(profit)

            # 库存
            # time.sleep(1)
            stock = tr.find_element_by_xpath('td[6]/div').get_attribute('innerHTML')
            stock_text.append(stock)

            info.append(
                {"lianjie": href_text[trs.index(tr)], "tupian": img_text[trs.index(tr)],
                 "biaoti": name_text[trs.index(tr)], "jiage": price_text[trs.index(tr)],
                 "lirun": profit_text[trs.index(tr)], "kucun": stock_text[trs.index(tr)]})

        # for next page
        driver.find_element_by_xpath('//*[@id="weidianHelp"]/div/div[3]/div[2]/div[2]/button[2]').click()
        time.sleep(3)
    print(info)
    print("接下来将数据传入Excel")
    return info

# 写入excel。
def xlsbook():
    xls = xlwt.Workbook()
    sheet = xls.add_sheet("当日库存情况")
    sheet.write(0, 0, "序号")
    sheet.write(0, 1, "商品链接")
    sheet.write(0, 2, "标题")
    sheet.write(0, 3, "价格")
    sheet.write(0, 4, "利润")
    sheet.write(0, 5, "库存")
    sheet.write(0, 6, "图片网址")
    for key, val in enumerate(info):
        sheet.write(key + 1, 0, "第" + str(key + 1) + "个")
        sheet.write(key + 1, 1, val.get("lianjie"))
        sheet.write(key + 1, 2, val.get("biaoti"))
        sheet.write(key + 1, 3, val.get('jiage'))
        sheet.write(key + 1, 4, val.get('lirun'))
        sheet.write(key + 1, 5, val.get('kucun'))
        sheet.write(key + 1, 6, val.get('tupian'))
    xls.save("d:/weidian-quwei/微店商品库存详细-" + a[0:4] + '年' + a[5:7] + '月' + a[8:10] + '日-总' + str(len(info)) + "个.xls")
    time.sleep(3)
    xls.save(r"D:\weidian-quwei\微店商品库存详细.xls")

# PART-TWO
# 读取素材--下载图片
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


# 下载图片函数
def download():
    root = "D:/weidian-quwei/微店商品图片2021/"
    for key, val in enumerate(info_duqu_xls):
        photo_dir = val.get("ziyuan")
        photo_biaoti = val.get("biaoti")
        path = root + photo_biaoti + '.jpg'

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


def look_kucun():
    parselweb()
    xlsbook()


def tupian_download():
    xls_duqu()
    download()


if __name__ == '__main__':
    info = []
    info_duqu_xls = []
    root = "D:/weidian-quwei/"
    if not os.path.exists(root):
        os.mkdir(root)
    else:
        print('提示：文件夹已存在')
    a = time.strftime("%Y-%m-%d %X", time.localtime())
    url = "https://d.weidian.com/weidian-pc/weidian-loader/#/pc-vue-fx-fx-item-manage/list"
    url1 = 'https://weidian.com/item.html?itemID=1942454952799849067235'

    # get_system_cookies()
    # 爬取数据存入excel
    # look_kucun()

    # 读取下载好的xls
    # tupian_download()

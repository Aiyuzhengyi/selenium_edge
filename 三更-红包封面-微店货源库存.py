import os
import re
import time
import requests
import xlrd
import xlwt
from selenium import webdriver

'''三更主要特点：增加webdriver如何获取Cookie，已经利用Cookie登陆。
这里使用到了官方文档：https://www.selenium.dev/zh-cn/documentation/webdriver/browser/cookies/中提到的添加Cookie和获取Cookies。
driver.add_cookie({"name": "test1", "value": "cookie1"}) 和 print(driver.get_cookies())。
Python中一切皆为对象，既然生成了一个driver，也可以作为参数传给其他函数调用哦。

2021年12月10日16:27:28

'''
# 1.【selenium模块】遥控win10自带的Edge浏览器打开指定网页（输入账号密码，根据逻辑进行操作）。
# 2.特点：无论网站有多少页，获取网页中总计多少条，利用正则表达式提取int型数字并取整，即翻页次数。处理如下：（num = re.findall('\d+', text)和 nub_new = int(num[0]) // 100。
# 3.特点：无论单个网页有多少数量，通过获取指定范围所有的tr标签。 for循环每个tr标签内，利用xpath索取tr标签内相应字段。（这里注意，存在一种情况：selenium提取不了标签文本。如果提取的元素文本为空，这是可能就是定位的
#     元素被隐藏了，即需要判断是否被隐藏。怎么解决？通过textContent, innerText, innerHTML等属性获取。我选择的是innerHTML。）函数完成后retur 列表，后续待用。
# 4.【xlwt模块】写入info数据,利用time模块依据当天日期存储指定路径。
# 5.为方便获取网站上所有的图片，还写了两个函数：读取素材 和 下载图片 。（注意的是，由于path中有'/'，需要清除以防保存报错。
def parselweb(url, info):
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get(url)
    # Adds the cookie into current browser context
    driver.add_cookie({'domain': '.weidian.com', 'expiry': 1639124025, 'httpOnly': False, 'name': '__spider__sessionid', 'path': '/', 'secure': False, 'value': 'bba3b36efe593732'})
    # Get all available cookies
    print(driver.get_cookies())
    # phone = driver.find_element_by_css_selector(
    #     '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div.user-telephone > div > div > div > input')
    # phone.send_keys('account')
    # key = driver.find_element_by_css_selector(
    #     '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div:nth-child(2) > div > div > input')
    # key.send_keys('password')
    # submit = driver.find_element_by_css_selector(
    #     '#app > div.content-wrapper > div > div > div.flex.login-container-content > div.login-wrapper > div.logo-info > form > div:nth-child(4) > div > button')
    # submit.click()
    # time.sleep(3)
    # shop_choose = driver.find_element_by_css_selector('#app > div.content-wrapper > div > div > div.bottom > div.item')
    # shop_choose.click()
    # time.sleep(3)
    driver.refresh()
    time.sleep(3)
    fenxiao = driver.find_element_by_css_selector(
        '#weidianMenu > div.v-common-menu-list > div.v-common-menu > div:nth-child(8) > div > a > div.v-common-menu-name')
    fenxiao.click()
    time.sleep(3)

    fenxiaoshang = driver.find_element_by_css_selector(
        '#v-common-second-nav-wrapper > div:nth-child(1) > div.v-common-three-nav-wrapper > div:nth-child(7) > div')
    fenxiaoshang.click()
    time.sleep(4)

    driver.find_element_by_css_selector(
        '#weidianHelp > div > div.card > div.fx-seller-item-table > div.el-pagination > span.el-pagination__sizes > div > div.el-input.el-input--mini.el-input--suffix > input').click()
    time.sleep(3)
    driver.find_element_by_css_selector(
        'body > div.el-select-dropdown.el-popper > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > ul > li:nth-child(5)').click()
    time.sleep(4)
    # 以上为驱动浏览器打开相应的网址，输入账号密码登陆。

    # 翻页次数设定。
    text = driver.find_element_by_css_selector(
        '#weidianHelp > div > div.card > div.fx-seller-item-table > div.el-pagination > span.el-pagination__total').text
    num = re.findall('\d+', text)
    nub_new = int(num[0]) // 100
    nub_new1 = nub_new + 1


    # 实现循环
    for i in range(nub_new1):
        # 局部变量单次循环后清空操作
        href_text = []
        img_text = []
        name_text = []
        price_text = []
        profit_text = []
        stock_text = []
        trs = driver.find_elements_by_xpath(
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
        time.sleep(5)
    print(info)
    print("接下来将数据传入Excel")
    return info

#写入excel。
def xlsbook(info):
    a = time.strftime("%Y-%m-%d %X", time.localtime())
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
    xls.save("d:/微店商品库存详细-" + a[0:4] + '年' + a[5:7] + '月' + a[8:10] + '日-总' + str(len(info)) + "个.xls")


# 读取素材
def xls_duqu(info_duqu_xls):
    xls_read = xlrd.open_workbook(r"D:\微店商品\微店商品库存详细-2021年12月09日-总762个.xls")
    xls_sheet = xls_read.sheet_by_name("当日库存情况")
    nrows = xls_sheet.nrows
    ncols = xls_sheet.ncols
    # time.sleep(3)  等设置待时间

    # 依次读取指定文件内容
    for r in range(1, nrows):
        middle = {}
        for c, key in zip(range(ncols), ["xuhao", "biaoti", "tupian"]):
            x = xls_sheet.cell(r, c).value
            middle[key] = x
        info_duqu_xls.append(middle)
    return info_duqu_xls


# 下载图片模块
def download(info_duqu_xls):
    root = "D://微店商品图片2021//"
    for key, val in enumerate(info_duqu_xls):
        photo_dir = val.get("tupian")
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


if __name__ == '__main__':
    info = []
    info_duqu_xls = []
    url = "https://d.weidian.com/weidian-pc/weidian-loader/#/pc-vue-fx-fx-item-manage/list"
    url1 = 'https://weidian.com/item.html?itemID=1942454952799849067235'
    # 爬取数据存入excel
    parselweb(url, info)
    xlsbook(info)

    # 读取下载好的xls
    # xls_duqu(info_duqu_xls)
    # download(info_duqu_xls)

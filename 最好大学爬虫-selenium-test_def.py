from selenium import webdriver
import xlwt
import os
import requests
333

# 1.【selenium模块】遥控win10自带的Edge浏览器打开指定网页；
# 2.查看页面共19页，撇去数字“0”，开始循环点击网页；
# 3.每次点击一个网页页码，则使用【css选择器】定位指定【字段】 并提取出来存入变量 school_info；
# 4.【xlwt模块】写入数据


# 驱动浏览器，传入2个参数
def parselweb(url, school_info):
    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get(url)
    school_logo = []
    school_name_cn = []
    school_name_en = []
    for i in range(20):
        if i == 0:  # 除去“0”
            pass
        else:
            driver.find_element_by_css_selector(".ant-pagination-item.ant-pagination-item-" + str(i)).click()
            name_cn_text = driver.find_elements_by_css_selector("div>.name-cn")
            name_en_text = driver.find_elements_by_css_selector("div>.name-en")
            logo_image = driver.find_elements_by_css_selector("div.logo > img")
            for ii in name_cn_text:
                school_name_cn.append(ii.text)
            for ii in name_en_text:
                school_name_en.append(ii.text)
            for ii in logo_image:
                school_logo.append(ii.get_attribute("src"))  # 使用.attribute 方法获取属性值。
    # 合并数据
    for i in range(len(school_name_cn)):
        school_info.append({"name-cn": school_name_cn[i], "name-en": school_name_en[i], "href_logo": school_logo[i]})

    # 格式化输出
    for i in school_info:
        print(i)
    print("接下来将数据传入Excel")
    return school_info


# 传入Excel
def xlsbook(school_info):
    xls = xlwt.Workbook()
    sheet = xls.add_sheet("最好大学排名")
    sheet.write(0, 0, "中国大学排名")
    sheet.write(0, 1, "大学中文名称")
    sheet.write(0, 2, "大学英文名称")
    sheet.write(0, 3, "大学logo")
    for key, val in enumerate(school_info):
        sheet.write(key + 1, 0, "第" + str(key + 1) + "位")
        sheet.write(key + 1, 1, val.get("name-cn"))
        sheet.write(key + 1, 2, val.get("name-en"))
        sheet.write(key + 1, 3, val.get("href_logo"))
    xls.save("d:/zuimeidaxue20211011.xls")


# 下载logo模块
def download(school_info):
    root = "D://最好大学logo3-完整//"
    for key, val in enumerate(school_info):
        path = root + "第" + str(key + 1) + "位  " + val.get("name-cn") + val.get("name-en") + ".png"
        try:
            if not os.path.exists(root):
                os.mkdir(root)
            if not os.path.exists(path):
                r = requests.get(val.get("href_logo"))
                with open(path, 'wb') as f:
                    f.write(r.content)
                    f.close()
                    print(path + '文件保存成功')
            else:
                print('文件已存在')
        except:
            print('爬取失败')


def main():
    school_info = []
    url = "https://www.shanghairanking.cn/rankings/bcur/2020"
    parselweb(url, school_info)
    xlsbook(school_info)


# download(school_info)


main()

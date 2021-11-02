import time
from selenium import webdriver
import xlwt
import re
from collections import namedtuple
import threading  # 导入线程库

# 创建全局变量local_account
local_account = threading.local()


# 多个线程分别接收对应的账号和密码，提供给webdriver_open()使用。由于产生的是全局变量？可跨函数使用。
def process_account(dri):
    local_account.account = dri.account
    local_account.key = dri.key
    local_account.count = dri.count
    local_account.rank = dri.rank
    webdriver_open()

#  打开浏览器，登陆账号密码，打开微信工单>待审核。
def webdriver_open():
    info_dh = []
    info_yj = []

    driver = webdriver.Edge(executable_path='msedgedriver.exe')
    driver.get(url)
    driver.find_element_by_css_selector("#loginId").send_keys(local_account.account)
    driver.find_element_by_css_selector("#credential").send_keys(local_account.key)
    driver.find_element_by_css_selector(".btn_login").click()
    time.sleep(2)
    driver.find_element_by_css_selector("#menuLi_M36 > a").click()  # 打开微信工单。
    time.sleep(2)
    driver.find_element_by_css_selector("#menuLi_M36_M3605").click()  # 打开待审核。
    time.sleep(2)

    # 记录当前待审核数量，提供给for循环语句使用。不同线程如何分配数量
    def daishenhe_number():
        daishenhe_count = driver.find_element_by_css_selector(
            "body > div.page_content > div > div > div > div.datagrid-wrap.panel-body > div.datagrid-pager.pagination > div.pagination-info").text
        dai1 = daishenhe_count.split(',')
        dai2 = re.findall('[0-9]*', dai1[1])
        current_REXcount = int(dai2[1]) - 1
        return current_REXcount

    def process_gongdan():
        daishenhe_number()
        DH = driver.find_element_by_css_selector(
            "#datagrid-row-r1-2-0 > td:nth-child(1) > div").text  # 读取第一行单号框的文本。
        info_dh.append(DH)
        driver.find_element_by_css_selector(
            "#datagrid-row-r1-2-0 > td:nth-child(10) > div > a").click()  # 单击办理按钮,和微信待办模块位置不一致。
        time.sleep(2)
        driver.switch_to.default_content()  # 办理页面弹出，由于办理页面属于新增框架，新框架ID=#doTask_frm。（暂时离开原框架#iframe_M3605)
        iframe_sh = driver.find_element_by_css_selector("#doTask_frm")  # 重新存储工单审核元素,iframe_sh。
        driver.switch_to.frame(iframe_sh)  # 切换到选择的审核iframe_sh。
        driver.find_element_by_css_selector(
            " div.tabs-header > div.tabs-wrap > ul > li:nth-child(2)").click()  # 切换该框架不同选项卡，答复意见在办理页面第三个选项卡，抽取最后一笔意见
        zhyj = driver.find_element_by_css_selector(
            "li:nth-child(2) > dl > dt > span:nth-child(2)").text  # 找到平台流程信息中最后一笔，抽取文本，作为办结意见，zhyj。
        zhyj1 = zhyj + "。"  # 防止提取空值。
        driver.find_element_by_css_selector(
            "div.tabs-header > div.tabs-wrap > ul > li.tabs-first").click()  # 切换回来,结案意见填写。
        driver.find_element_by_css_selector("td>#comm").send_keys(zhyj1)  # 填写办结意见
        driver.find_element_by_css_selector(" tr > td > a:nth-child(3)").click()  # 审核通过按钮,按钮位置可能会变 注意！
        time.sleep(1)
        driver.find_element_by_css_selector(
            " div.dialog-button.messager-button > a:nth-child(1) ").click()  # 审核通过二次确定按钮
        info_yj.append(zhyj1)
        driver.find_element_by_css_selector('body > div > div > div > div.btn_area_setc > a').click()  # 点击关闭按钮，刷新循环列表，开启下一个循环。（此处无响应也可关闭办理页面）
        time.sleep(2)
        driver.switch_to.default_content()  # 框架焦点丢失，使用WebElement 重新锁定，#iframe_M3605。
        iframe = driver.find_element_by_css_selector("#iframe_M3605")
        driver.switch_to.frame(iframe)  # 切换到选择的iframe
        time.sleep(2)
        info.append({"danhao": info_dh[local_account.count], "yijian": info_yj[local_account.count]})
        print("序号：{0:<5}工单单号{1:^20}结案意见{2:^50}操作时间{3:>}".format(local_account.count + 1, info_dh[local_account.count],info_yj[local_account.count], time.ctime()))
        local_account.count += 1

# 循环配置
    iframe = driver.find_element_by_css_selector("#iframe_M3605")
    driver.switch_to.frame(iframe)  # 切换到选择的iframe-m3605 查询框架ID。
    # driver.find_element_by_css_selector('#claimStatus > option:nth-child(3)').click()  # 选择签收状态为“已签收”
    driver.find_element_by_css_selector('#order > option:nth-child(' + str(local_account.rank) + ')').click()
    time.sleep(1)
    driver.find_element_by_css_selector(
        " tbody > tr:nth-child(6) > td > a:nth-child(1)").click()  # 点击查询按钮，刷新页面,设置等待时间4s
    time.sleep(4)
    daishenhe_count = driver.find_element_by_css_selector(
        "body > div.page_content > div > div > div > div.datagrid-wrap.panel-body > div.datagrid-pager.pagination > div.pagination-info").text
    print("{0:=^80}".format("当前时间：" + time.ctime() + '   ' + daishenhe_count))
    while True:
        if daishenhe_number() > 0:
            process_gongdan()
            continue
        else:
            print('process_gongdan is over!')
            break
    # driver.quit()
    return info


# 写入excel。
def xls(info):
    print("新建Excel文件，正在写入D盘ing")
    xls = xlwt.Workbook()
    sheet = xls.add_sheet("审核明细")
    sheet.write(0, 0, "序号")
    sheet.write(0, 1, "工单编号")
    sheet.write(0, 2, "答复意见")

    for key, val in enumerate(info):
        sheet.write(key + 1, 0, "第" + str(key + 1) + "张")
        sheet.write(key + 1, 1, val.get("danhao"))
        sheet.write(key + 1, 2, val.get("yijian"))

    xls.save("d:/" + info[0].get("danhao") + ".xls")


if __name__ == '__main__':
    info = []
    url = "http://网站地址"
    drivers = namedtuple('drivers', 'account key count rank')
    list_driver = [drivers(账号。)]
    t1 = threading.Thread(target=process_account, args=(list_driver[0],), name='Thread-A11014')
    t2 = threading.Thread(target=process_account, args=(list_driver[1],), name='Thread-B11014')
    t1.start()
    t2.start()


    # xls(info)

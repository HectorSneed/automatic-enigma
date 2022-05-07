import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import xlwings as xw
import requests
import time
from lxml import etree


def openChrome():
    opt = webdriver.ChromeOptions()
    opt.add_argument("--start-maximized")
    opt.add_argument("user-data-dir=chrome/Data/")
    opt.add_argument("--remote-debugging-port=9222")
    opt.binary_location = "chrome/ChromeFP.exe"
    browser = webdriver.Chrome(executable_path="chrome/chromedriver.exe", options=opt)
    return browser


def openExcel(file_name):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = True
    wb = xw.Book(file_name)
    sht = wb.sheets.active
    return sht


def analyExcel(sht, pos_input, begin_pos_int, end_pos_int):
    values = []
    begin_pos = pos_input + str(begin_pos_int)
    end_pos = pos_input + str(end_pos_int)
    value = sht.range(begin_pos + ':' + end_pos).value
    if isinstance(value, list):
        return value
    elif isinstance(value, str):
        values.append(value)
        return values
    else:
        return "PASS"


def getAddrByCompanyName(name, browser, dir_path):
    url = "https://www.tianyancha.com/search?key=" + name
    r = requests.get(url)
    if r.status_code != 200:
        return "查询失败"
    html = etree.HTML(r.text)
    href = html.xpath('//div[contains(@class, "search-result-single")]//div[@class="contact row"]//span[last()]/text()')
    if len(href) == 0:
        return "查询失败"
    else:
        browser.get(url)
        time.sleep(2)
        try:
            tmp = browser.find_element_by_xpath('//div[@class="expand toggle-btn"]')
        except NoSuchElementException as msg:
            return "查询失败"
        else:
            flag = tmp.text == "收起"
            if flag:
                tmp.click()
            dir_path = dir_path + '\\' + name + "_天眼查.png"
            browser.get_screenshot_as_file(dir_path)
            return href[0]


def getDistenceByAmap(cur_input_addr, cur_output_addr, browser, dir_path):
    url = "https://www.amap.com/dir"
    browser.get(url)
    time.sleep(3)
    browser.find_element_by_css_selector('#dir_from_ipt').send_keys(cur_input_addr)
    browser.find_element_by_css_selector('#dir_to_ipt').send_keys(cur_output_addr)
    browser.find_element_by_css_selector('.dir_submit').click()
    time.sleep(2)
    try:
        browser.find_element_by_xpath('//div[@class="choose-poi-content" and @dirtype="from"]')
    except NoSuchElementException as msg:
        print("no from list")
    else:
        try:
            browser.find_element_by_xpath(
                '//div[@class="choose-poi-content" and @dirtype="from"]//li[contains(@class, choose_0)]').click()
        except NoSuchElementException as msg:
            return "路径规划失败，无法选取源地址"
    time.sleep(2)

    try:
        browser.find_element_by_xpath('//div[@class="choose-poi-content" and @dirtype="to"]')
    except NoSuchElementException as msg:
        print("no to list")
    else:
        try:
            browser.find_element_by_xpath(
                '//div[@class="choose-poi-content" and @dirtype="to"]//li[contains(@class, choose_0)]').click()
        except NoSuchElementException as msg:
            return "路径规划失败，无法选取目的地址"

    # browser.find_element_by_css_selector('.dir_submit').click()
    time.sleep(2)
    distence = browser.find_element_by_xpath('//*[@id="plantitle_0"]/p/span[2]').text
    return distence


def getDistenceByBaidu(cur_input_addr, cur_output_addr, browser, dir_path, cur_input_name):
    url = "https://map.baidu.com"
    browser.get(url)
    time.sleep(3)
    try:
        close_btn = browser.find_element_by_xpath(
            '//div[@id="passport-login-pop"]//div[@class="buttons"]//a[@class="close-btn"]')
    except NoSuchElementException as msg:
        pass
    else:
        close_btn.click()
    try:
        path_btn = browser.find_element_by_xpath('//*[@id="sole-searchbox-content"]/div[2]')
    except NoSuchElementException as msg:
        return "查询失败1"
    path_btn.click()
    time.sleep(1)
    try:
        drive_btn = browser.find_element_by_xpath('//*[@id="route-searchbox-content"]/div[1]/div[1]/div[2]')
    except NoSuchElementException as msg:
        return "查询失败2"
    drive_btn.click()
    try:
        start_input = browser.find_element_by_xpath('//*[@id="route-searchbox-content"]/div[2]/div/div[2]/div[1]/input')
    except NoSuchElementException as msg:
        return "查询失败3"
    start_input.send_keys(cur_input_addr)
    time.sleep(1)
    try:
        end_input = browser.find_element_by_xpath('//*[@id="route-searchbox-content"]/div[2]/div/div[2]/div[2]/input')
    except NoSuchElementException as msg:
        return "查询失败4"
    end_input.send_keys(cur_output_addr)
    time.sleep(1)
    try:
        search_btn = browser.find_element_by_xpath('//*[@id="search-button"]')
    except NoSuchElementException as msg:
        return "查询失败5"
    search_btn.click()
    time.sleep(3)
    try:
        choose_start = browser.find_element_by_xpath('//*[@id="RA_ResItem_0"]/table/tbody/tr[1]/td[2]/div')
    except NoSuchElementException as msg:
        return "查询失败6"
    choose_start.click()
    time.sleep(1)
    try:
        choose_end = browser.find_element_by_xpath('//*[@id="RA_ResItem_1"]/table/tbody/tr[1]/td[2]/div')
    except NoSuchElementException as msg:
        return "查询失败7"
    choose_end.click()
    time.sleep(1)
    try:
        distence = browser.find_element_by_xpath('//*[@id="navtrans_content"]/div[1]/div[1]/div[1]/p[1]/span[2]')
    except NoSuchElementException as msg:
        return "查询失败8"
    dir_path = dir_path + '\\' + cur_input_name + "_百度地图.png"
    browser.get_screenshot_as_file(dir_path)
    return distence.text


def solution(sht, name_input, addr_input, addr_output_pos, distence_output_pos, begin_pos_int, browser, dir_path):
    print(name_input)
    print(addr_input)
    numbers = len(name_input)
    cur_num = 0
    while cur_num < numbers:
        cur_input_name = name_input[cur_num]
        print(cur_input_name)
        cur_input_addr = addr_input[cur_num]
        print(cur_input_addr)
        if cur_input_name == "PASS":
            cur_num += 1
            begin_pos_int += 1
            continue
        cur_output_addr = getAddrByCompanyName(cur_input_name, browser, dir_path)
        print(cur_output_addr)
        cur_output_addr_pos = addr_output_pos + str(begin_pos_int)
        print(cur_output_addr_pos)
        sht.range(cur_output_addr_pos).value = cur_output_addr

        cur_output_dis_pos = distence_output_pos + str(begin_pos_int)
        print(cur_output_dis_pos)
        cur_output_dis = getDistenceByBaidu(cur_input_addr, cur_output_addr, browser, dir_path, cur_input_name)
        print(cur_output_dis)
        sht.range(cur_output_dis_pos).value = cur_output_dis
        cur_num += 1
        begin_pos_int += 1


def analyFile(file_name, dir_path, begin_pos_int, end_pos_int, name_input_pos, addr_input_pos, \
              addr_output_pos, distence_output_pos):
    print(begin_pos_int)
    print(end_pos_int)
    sht = openExcel(file_name)
    time.sleep(1)
    browser = openChrome()
    time.sleep(1)
    name_input = analyExcel(sht, name_input_pos, begin_pos_int, end_pos_int)
    addr_input = analyExcel(sht, addr_input_pos, begin_pos_int, end_pos_int)
    solution(sht, name_input, addr_input, addr_output_pos, distence_output_pos, begin_pos_int, browser, dir_path)

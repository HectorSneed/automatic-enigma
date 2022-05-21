import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException

import xlwings as xw
import time

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
    return sht, wb


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


def findAddrTmp(i, browser):
    partxpath = '//*[@id="page-container"]/div/div[2]/section/main/div[' + str(i) + ']/div[2]/div[1]/div/div[2]/div[2]'
    j = 2
    while j < 6:
        labelxpath = partxpath + '/div[' + str(j) + ']/div/span[1]'
        try:
            labeltmp = browser.find_element_by_xpath(labelxpath)
        except NoSuchElementException as msg:
            pass
        else:
            if labeltmp.text == "地址：":
                addrxpath = partxpath + '/div[' + str(j) + ']/div/span[2]'
                try:
                    addrtmp = browser.find_element_by_xpath(addrxpath)
                except NoSuchElementException as msg:
                    pass
                else:
                    return addrtmp
        j += 1
    return None


def checkNameXpath(browser, name):
    i = 2
    while i < 4:
        namexpath = '//*[@id="page-container"]/div/div[2]/section/main/div[' + str(i) + ']/div[2]/div[1]/div/div[2]/div[2]/div[1]/div[1]/a/span'
        try:
            nametmp = browser.find_element_by_xpath(namexpath)
        except NoSuchElementException as msg:
            pass
        else:
            if nametmp.text == name:
                return findAddrTmp(i, browser)
        i += 1
    return None


def checkTmpBtn(browser):
    i = 1
    while i < 3:
        btnxpath = '//*[@id="page-container"]/div/div[2]/section/main/div[' + str(i) + ']/div/div/div[3]/div'
        try:
            btntmp = browser.find_element_by_xpath(btnxpath)
        except NoSuchElementException as msg:
            pass
        else:
            if btntmp.is_displayed() and btntmp.is_enabled() and (btntmp.text == "展开" or btntmp.text == "收起"):
                return btntmp
        i += 1
    return None


def getAddrByCompanyName(name, browser, dir_path):
    url = "https://www.tianyancha.com/search?key=" + name
    browser.get(url)
    time.sleep(2)
    flag = True
    count = 0
    while flag and count < 10:
        addr = checkNameXpath(browser, name)
        if addr is None:
            count += 1
            browser.refresh()
            time.sleep(1)
            continue

        tmpbtn = checkTmpBtn(browser)
        if tmpbtn is None:
            count += 1
            browser.refresh()
            time.sleep(1)
            continue
        if tmpbtn.text == "收起":
            try:
                tmpbtn.click()
            except ElementClickInterceptedException as msg:
                count += 1
                browser.refresh()
                time.sleep(1)
                continue
        flag = False
    if flag:
        return "查询失败0"
    dir_path = dir_path + '\\' + name + "_天眼查.png"
    browser.get_screenshot_as_file(dir_path)
    return addr.text


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
    timeout = WebDriverWait(browser, 5)
    time.sleep(3)
    try:
        close_btn = browser.find_element_by_xpath(
            '//div[@id="passport-login-pop"]//div[@class="buttons"]//a[@class="close-btn"]')
    except NoSuchElementException as msg:
        pass
    else:
        close_btn.click()
    flag = True
    count = 0
    while flag and count < 10:
        try:
            path_btn = timeout.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sole-searchbox-content"]/div[2]')))
        except TimeoutException as msg:
            pass
        else:
            path_btn.click()
        try:
            drive_btn = timeout.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="route-searchbox-content"]/div[1]/div[1]/div[2]')))
        except TimeoutException as msg:
            pass
        else:
            drive_btn.click()
        try:
            start_input = timeout.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="route-searchbox-content"]/div[2]/div/div[2]/div[1]/input')))
        except TimeoutException as msg:
            pass
        else:
            start_input.send_keys(cur_input_addr)
        try:
            end_input = timeout.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="route-searchbox-content"]/div[2]/div/div[2]/div[2]/input')))
        except TimeoutException as msg:
            pass
        else:
            end_input.send_keys(cur_output_addr)
            time.sleep(2)
        try:
            search_btn = timeout.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="search-button"]')))
        except TimeoutException as msg:
            pass
        else:
            search_btn.click()

        try:
            label_same = browser.find_element_by_xpath('//*[@id="toast-wrapper"]')
        except NoSuchElementException as msg:
            pass
        else:
            if label_same.is_displayed():
                return "0米"

        try:
            choose_start = timeout.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="RA_ResItem_0"]/table/tbody/tr[1]/td[2]/div')))
        except TimeoutException as msg:
            pass
        else:
            if choose_start.is_displayed():
                choose_start.click()

        try:
            choose_end = timeout.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="RA_ResItem_1"]/table/tbody/tr[1]/td[2]/div')))
        except TimeoutException as msg:
            pass
        else:
            if choose_end.is_displayed():
                choose_end.click()

        try:
            distence = timeout.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="navtrans_content"]/div[1]/div[1]/div[1]/p[1]/span[2]')))
        except TimeoutException as msg:
            count += 1
            browser.refresh()
            continue
        flag = False
    dir_path = dir_path + '\\' + cur_input_name + "_百度地图.png"
    browser.get_screenshot_as_file(dir_path)
    return distence.text


def solution(sht, name_input, addr_input, addr_output_pos, distence_output_pos, begin_pos_int, browser, dir_path):
    #print(name_input)
    #print(addr_input)
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
        sht.range(cur_output_addr_pos).value = cur_output_addr

        cur_output_dis_pos = distence_output_pos + str(begin_pos_int)
        if cur_input_addr == cur_output_addr:
            cur_output_dis = "0米"
        else:
            cur_output_dis = getDistenceByBaidu(cur_input_addr, cur_output_addr, browser, dir_path, cur_input_name)
        print(cur_output_dis)
        sht.range(cur_output_dis_pos).value = cur_output_dis
        cur_num += 1
        begin_pos_int += 1


def analyFile(file_name, dir_path, begin_pos_int, end_pos_int, name_input_pos, addr_input_pos, \
              addr_output_pos, distence_output_pos):
    sht, wb = openExcel(file_name)
    time.sleep(1)
    browser = openChrome()
    time.sleep(1)
    name_input = analyExcel(sht, name_input_pos, begin_pos_int, end_pos_int)
    addr_input = analyExcel(sht, addr_input_pos, begin_pos_int, end_pos_int)
    solution(sht, name_input, addr_input, addr_output_pos, distence_output_pos, begin_pos_int, browser, dir_path)
    wb.save()

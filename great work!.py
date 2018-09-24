import xlrd
from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
data=xlrd.open_workbook("ticket_catalog.xlsx")
table=data.sheets()[0]
nrows=table.nrows

drivers = webdriver.Chrome()  # 创建一个edge浏览器对象
drivers.get("http://www.baiwang.com/cterminal/cterminal/fpcy.html")
drivers.find_element_by_class_name("money").click()
a=input()
drivers.find_element_by_id("in_yzm").send_keys(a)
element1 = drivers.find_element_by_id("in_bdbh")
element2 = drivers.find_element_by_id("in_fpje")
for i in range(1,nrows):
    time.sleep(2)
    ActionChains(drivers).click(element1).send_keys(table.cell(i,0).value).perform()
    #drivers.find_element_by_xpath("//input[@id='in_bdbh']").send_keys(table.cell(i,0).value)
    ActionChains(drivers).click(element2).send_keys(table.cell(i,1).value).perform()
    #drivers.find_element_by_id("in_fpje").send_keys(table.cell(i,1).value)
    drivers.find_element_by_id("btn_fpcy").click()
    time.sleep(2)
    drivers.find_element_by_xpath("//div[@id='cboxLoadedContent']//a[@id='btn_bwsjdy']").click()
    element3 = drivers.find_element_by_css_selector("#cboxClose")
    ActionChains(drivers).click(element3).click(element3).perform()
    drivers.find_element_by_id("in_bdbh").clear()
    drivers.find_element_by_id("in_fpje").clear()

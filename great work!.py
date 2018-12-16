import xlrd
import os
import shutil
from selenium import webdriver
import time
import requests
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
data=xlrd.open_workbook("ticket_catalog.xlsx")
table=data.sheets()[0]
nrows=table.nrows

filename=table.cell(0,2).value;
#print (filename)
os.mkdir(filename)

name=table.cell(0,0).value[4:10]
#print(name)
os.makedirs(filename+"\\"+name)


drivers = webdriver.Chrome()  # 创建一个Chrome浏览器对象
drivers.get("http://www.baiwang.com/cterminal/cterminal/fpcy.html")
time.sleep(2)

drivers.find_element_by_class_name("money").click()
a=input()
drivers.find_element_by_id("in_yzm").send_keys(a)
element1 = drivers.find_element_by_id("in_bdbh")
element2 = drivers.find_element_by_id("in_fpje")

for i in range(0,nrows):
    time.sleep(1)
    u_name=table.cell(i,0).value[4:10]
    if(u_name!=name):
       time.sleep(1)
       for root, dirs, files in os.walk("C:\\Users\\NAME\\Downloads"):
           for file in files:
               shutil.move("C:\\Users\\NAME\\Downloads\\"+file,filename+"\\"+name)
       name=u_name
       try:
          os.makedirs(filename+"\\"+name)
       except:
          print(str(filename)+"\\"+str(name)+" 已经存在")
    try: 
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
    except NoSuchElementException:
        i=i+1
        try:
            ActionChains(drivers).click(element3).click(element3).perform()
        except:
            print("not open")
        drivers.find_element_by_id("in_bdbh").clear()
        drivers.find_element_by_id("in_fpje").clear()
        time.sleep(2)
        print("Lost one ticket!")
time.sleep(1)
for root, dirs, files in os.walk("C:\\Users\\NAME\\Downloads"):
    for file in files:
      shutil.move("C:\\Users\\NAME\\Downloads\\"+file,filename+"\\"+name)

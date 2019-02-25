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

choice=input("请输入模式（A：默认模式，B：选择模式）：")
if choice=='B':
    lo=input("请输入下限：")
    hi=input("请输入上限：")
    point_choice=input("是否需要特殊小数点（A：需要，B：不需要）：")
elif choice=='A':
    lo=0
    hi=999999
    point_choice='B'

drivers.find_element_by_class_name("money").click()
a=input("输入验证码：")
drivers.find_element_by_id("in_yzm").send_keys(a)
element1 = drivers.find_element_by_id("in_bdbh")
element2 = drivers.find_element_by_id("in_fpje")

route="C:\\Users\\NAME\\Downloads"

def search(bdbh,fpje):
    try: 
         ActionChains(drivers).click(element1).send_keys(bdbh).perform()
         #drivers.find_element_by_xpath("//input[@id='in_bdbh']").send_keys(table.cell(i,0).value)
         ActionChains(drivers).click(element2).send_keys(fpje).perform()
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

def ispoint(fpje,point_choice):
    if point_choice=='B':
        return True
    elif point_choice=='A':
        if fpje[::-1][0:2] != "00":
            return True
        else:
            return False


for i in range(0,nrows):
    time.sleep(1)
    u_name=table.cell(i,0).value[4:10]
    if(u_name!=name):
       time.sleep(1)
       for root, dirs, files in os.walk(route):
           for file in files:
               shutil.move(route+"\\"+file,filename+"\\"+name)
       name=u_name
       try:
          os.makedirs(filename+"\\"+name)
       except:
          print(str(filename)+"\\"+str(name)+" 已经存在")
    if float(lo)<=float(table.cell(i,1).value) and float(hi)>=float(table.cell(i,1).value) and ispoint(table.cell(i,1).value,point_choice):
        search(table.cell(i,0).value,table.cell(i,1).value)
        time.sleep(1)

for root, dirs, files in os.walk(route):
    for file in files:
      shutil.move(route+"\\"+file,filename+"\\"+name)


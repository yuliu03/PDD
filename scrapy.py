import re
import sys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from openpyxl import load_workbook, Workbook
from selenium.webdriver import Chrome
from selenium.webdriver import ChromeOptions


#总览 //*[@id="root"]/div/div/nav/div/div[4]/ul/li[1]
overView = "//*[@id='root']/div/div/nav/div/div[4]/ul/li[1]"
#交易数据
dealInfo = "//*[@id=root']/div/div/nav/div/div[6]/ul/li[3]/a/span"
#流量数据
flow = "//*[@id='root']/div/div/nav/div/div[6]/ul/li[5]/a/span"
#商品数据
productInfo = "//*[@id='root']/div/div/nav/div/div[6]/ul/li[2]/a/span"
#服务数据
serviceInfo = "//*[@id='root']/div/div/nav/div/div[6]/ul/li[4]/a/span"


chromedriver = "D:\Google\Chrome\Application\chromedriver.exe"  # 这里是你的驱动的绝对地址

#总览页面
def runOverView(driver):
    print("总览页面")
    btn = driver.find_element_by_xpath((By.XPATH, overView))
    print(btn.get_attribute('innerHTML'))
    btn.click()
    time.sleep(1)

#交易数据页面
def runDealInfo(driver):
    print("交易数据页面")
    btn = driver.find_element_by_xpath((By.XPATH, dealInfo))
    btn.click()
    print(btn.get_attribute('innerHTML'))
    time.sleep(1)


#流量数据页面
def runFlowInfo(driver):
    print("流量数据页面")
    btn = driver.find_element_by_xpath((By.XPATH, flow))
    print(btn.get_attribute('innerHTML'))

#商品数据页面
def runProductInfo(driver):
    print("商品数据页面")
    btn = driver.find_element_by_xpath((By.XPATH, productInfo))
    btn.click()
    print(btn.get_attribute('innerHTML'))
    time.sleep(1)


#服务数据页面
def runServiceInfo(driver):
    print("服务数据页面")
    btn = driver.find_element_by_xpath((By.XPATH, serviceInfo))
    btn.click()
    print(btn.get_attribute('innerHTML'))
    time.sleep(1)



option = ChromeOptions()
option.add_experimental_option('excludeSwitches', ['enable-automation'])

driver = webdriver.Chrome(executable_path=chromedriver,options=option)

waitTime = 5

driver.get('https://mms.pinduoduo.com/home/')
wait = WebDriverWait(driver, waitTime)

#登录方式按钮
loginChooseId = "//*[@id='root']/div/div/div/main/div/section[2]/div/div[1]/div[1]/div/div[2]"
userId = "//*[@id='usernameId']"
password = "//*[@id='passwordId']"
loginBtnId = "//*[@id='root']/div/div/div/main/div/section[2]/div/div[1]/div[2]/section/div/div[2]/button"

wait.until(EC.presence_of_element_located((By.XPATH, loginChooseId)))
loginChooseBtn = driver.find_element_by_xpath(loginChooseId)
loginChooseBtn.click()


wait = WebDriverWait(driver, waitTime)
wait.until(EC.presence_of_element_located((By.XPATH, userId)))
userName = driver.find_element_by_xpath(userId)
userName.send_keys("物产数码_刘愉")

wait.until(EC.presence_of_element_located((By.XPATH, password)))
userName = driver.find_element_by_xpath(password)
userName.send_keys("Ly123456")

#submitBtn
time.sleep(0.5)
wait.until(EC.presence_of_element_located((By.XPATH, loginBtnId)))
login = driver.find_element_by_xpath(loginBtnId)
login.click()


#/html/body/div[13]/div/div/div[2]/div/div[2]/button/span

input("enter when is ok ")
driver.get('https://mms.pinduoduo.com/sycm/evaluation/overview')
rootElem = driver.find_element_by_xpath((By.XPATH,"//*[@id='root']"))
print(rootElem.get_attribute('innerHTML'))

#/html/body/div[13]/div/div/div[2]/div/div[2]/button/span
nextBtn = "/html/body/div[13]/div/div/div[2]/div/div[2]/button/span"
wait.until(EC.presence_of_element_located((By.XPATH, nextBtn)))
next = driver.find_element_by_xpath(nextBtn)
next.click()

# runDealInfo(rootElem)
# input("enter to next ")
# runFlowInfo(rootElem)
# input("enter to next ")
# runOverView(rootElem)
# input("enter to next ")
# runProductInfo(rootElem)
# input("enter to next ")
# runServiceInfo(rootElem)
# input("enter to next ")

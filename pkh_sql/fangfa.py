
from selenium.common.exceptions import ElementNotVisibleException,NoSuchElementException,TimeoutException,StaleElementReferenceException
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time

def retryingFindClick(by,driver):
    result = False
    attempts = 0
    while attempts < 2:
        try:
            driver.find_element_by_xpath(by).click()
            result = True
            break
        except StaleElementReferenceException as e :
            pass
        attempts+=1
        return result

def input_driver (driver,value,xpath):
    if value==None:
        pass
    else:
        driver.find_element_by_xpath(xpath).clear()
        driver.find_element_by_xpath(xpath).send_keys(value)

def select_driver (driver,value,xpath):

    if value==None or driver.find_element_by_xpath(xpath+"/../../label").text==value:
        pass
    else:        
        driver.find_element_by_xpath(xpath).click()
        time.sleep(0.3)
        class_name =driver.find_element_by_xpath(xpath+'/../..').get_attribute("class") 
        driver.find_element_by_xpath("//span[@class='%s'][contains(text(),'%s')]" % (class_name[:12],value)).click()
        time.sleep(0.3)

def radio_driver(driver,value,xpath):
    if value==None:
        pass
    else:
        if "circle" in driver.find_element_by_xpath(xpath.replace('?',value)).get_attribute("class"):
            pass
        else:
            driver.find_element_by_xpath(xpath.replace('?',value)).click()

def radio_driver_zp(driver,value,xpath):
    if value==None or value=="":
        pass
    else:
        if "circle" in driver.find_element_by_xpath(xpath.replace('?',value)).get_attribute("class"):
            pass
        elif   "circle" in driver.find_element_by_xpath(("//radio-checkbox-code[@groupname='aac008_2']/*/*/div[?]/*/*/div[2]/span").replace('?',value)).get_attribute("class"):
            driver.find_element_by_xpath("//radio-checkbox-code[@groupname='aac008_2']/*/*/div[?]/*/*/div[2]/span").replace('?',value).click()
            driver.find_element_by_xpath(xpath.replace('?',value)).click()
        elif   "circle" in driver.find_element_by_xpath(("//radio-checkbox-code[@groupname='aac008_3']/*/*/div[?]/*/*/div[2]/span").replace('?',value)).get_attribute("class"):
            driver.find_element_by_xpath("//radio-checkbox-code[@groupname='aac008_3']/*/*/div[?]/*/*/div[2]/span").replace('?',value).click()
            driver.find_element_by_xpath(xpath.replace('?',value)).click()
        elif   "circle" in driver.find_element_by_xpath(("//radio-checkbox-code[@groupname='aac007']/*/*/div[?]/*/*/div[2]/span").replace('?',value)).get_attribute("class"):
            driver.find_element_by_xpath("//radio-checkbox-code[@groupname='aac007']/*/*/div[?]/*/*/div[2]/span").replace('?',value).click()
            driver.find_element_by_xpath(xpath.replace('?',value)).click()
        else:
            driver.find_element_by_xpath(xpath.replace('?',value)).click()

def p_mselect_driver(driver,value,xpath):
    if value==None:
        pass
    else:
        a=sorted(value.split(","))

        for i in driver.find_element_by_xpath(xpath+"/../../div[2]").get_attribute("title").split(',').sort():
            driver.find_element_by_xpath(xpath+("/../../div[4]/div[2]//label[text()='%s']"% i)).click()
        for j in a:
            driver.find_element_by_xpath(xpath+("/../../div[4]/div[2]//label[text()='%s']"% j)).click()
        
# coding:utf-8
# include <sys.stdin>
# include <sys.stdout>
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import os
import time
import urllib
import urllib3
import fangfa as ff
import exceltojson as e_to_j
import personData
import setting
import json
import sys
import sqlite3
import sqlff
from selenium.common.exceptions import ElementNotVisibleException, NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
from selenium.webdriver.common.by import By

from datetime import datetime
import pytesser3 as pyt3
from PIL import Image,ImageEnhance
from io import BytesIO


#e_to_j.js_to_xlsx('002.json','error_%s.xlsx' % str(datetime.now().strftime("%Y-%m-%d")))


# 读取json数据库,获取待处理列表
list_poor_family_js = []  # 所有待处理的贫困户数据JS类型
list_poor_family = []  # 所有待处理的贫困户数据
familyDatas = []
n = 0  # 计数器
end = 0
start = 0
error = 0
switch_1 = True
error_count = 0
time_clock = 5
b = True
my = setting.setting("","")
js_below = "document.querySelector('body > div.swal2-container.swal2-fade.swal2-shown > div > div.swal2-buttonswrapper > button.swal2-confirm.swal2-styled').click()"

#while True:
#    print('''
######################################################################################
#########################贫困户结对帮扶责任人自动维护脚本 V1.2########################
##################################### 作者：小王 #####################################
################################### QQ:282028623 #####################################
##################################### 2018-8-24 ######################################
######################################################################################
#说明：本脚本仅供学习交流，造成后果由使用者自行负责！！
#--------------------------------------------------------------------------------
#准备工作：
#    1.将所有结对帮扶人录入到指定结对帮扶单位，并将各结对帮扶单位与被帮扶村完成关联；
#    2.尽量去除重复的结对帮扶人，保证按姓名查询结对帮扶人结果唯一，这样确保默认点选的第一个就是正确的结对帮扶人；
#    3.同名的结对帮扶人，请在添加结对帮扶人加入空格，以区分，避免录入交叉。
#    4.请将与Chrome浏览器版本对应的chromedriver，并放在Chrome程序目录Application文件夹中；
#    5.首次运行，请按照脚本目录下excel模板编写数据源，命名为"f.xlsx"和"p.xlsx"后放在脚本目录；
#    6.路径下的"001.xlsx","001.json","002.json"为数据转换文件，请勿删除；
#    7.第一次运行：请按照excel模板在目录下放入excel文件，并命名为"001.xlsx";
#    8.非首次运行，请确保上次保存的"002.json"存在;

#    操作过程中请勿点击页面和最小化，可以在后台运行

#                       首次运行或者新的数据源请输入  1
#                        利用现存的数据源运行请输入   2
#                       导出现存数据中错误数据到excel 3
 
#    ''')

#    input_num = input("请输入：(1/2/3)  >>>>>>>>>")
#    if input_num == "1":
#        my.get_and_save_setting()
#        e_to_j.excel_json('./file/f.xlsx' or './file/f.xls', './file/f.json')
#        e_to_j.excel_json('./file/p.xlsx' or './file/p.xls', './file/p.json')
#        e_to_j.file_to_json_fomat(
#            "./file/f.json", "./file/p.json","./file/fp.json", list_poor_family)
#        # e_to_j.json_to_personDatalist("002.json",list_poor_family_js,list_poor_family,error,start,end,n)
#        input("数据转换完成，生成文件完成，程序即将关闭，下次使用请输入2，按任意键退出......")
#        sys.exit()

#    elif input_num == "2":
#        my.load_setting()

#        e_to_j.json_to_familyDatalist(
#            "./file/fp.json", list_poor_family_js, list_poor_family, error, start, end, n)
#        break
#    elif input_num == "3":
#        e_to_j.js_to_xlsx('./file/fp.json', './file/error.xlsx')
#        # 将002.json中error=true的项提取出来并写入error+datetime.excel
#        input("按任意键退出......")
#        sys.exit()
#    elif input_num == "4":
#        e_to_j.js_to_xlsx_all('./file/fp.json', './file/all.xlsx')

#        input("按任意键退出......")
#        sys.exit()
#    elif input_num =="5":
#        break
#    else:
#        print("请重新输入")
# input()

#e_to_j.personDatalist_to_json(list_poor_family, './file/fp_on.json')

# 初始化---------
# input()
# e_to_j.save_as_json(list_poor_family,'002.json')

# 构造模拟浏览器
cnn =sqlite3.connect("./file/test.db")
csor =cnn.cursor()#创建指针
familynums = sqlff.loadfamilynums(csor) 

#my.get_and_save_setting()
my.load_setting()



chromedriver = my.chromePath_old

os.environ["webdriver.chrome.driver"] = chromedriver

driver = webdriver.Chrome(chromedriver)  # 模拟打开浏览器
# driver.implicitly_wait(time_clock)

# WebDriverWait(driver, 10).until(EC.invisibility_of_element_located(By.XPATH,
# my.xpath11))


url = my.url

driver.get(url)  # 打开网址
driver.maximize_window()  # 窗口最大化
# time.sleep(8) #等待网页加载20秒

#！！尚未加载 异常处理

# driver.find_element_by_xpath('//*[@id="u1"]/a[7]').click() #点击登录按钮

# time.sleep(2)
# driver.find_element_by_class_name('ng-pristine ng-invalid ui-inputtext
# ui-corner-all ui-state-default ui-widget ng-touched').clear() #先清除输入框内容
# driver.find_element_by_xpath("//input[@formcontrolname='account']").send_keys(u'43072320801')
# 输入账号
time.sleep(2)


driver.find_element_by_xpath(my.xpath1).send_keys(my.account)  # 输入账号
# time.sleep(1)
driver.find_element_by_xpath(my.xpath2).send_keys(my.password)  # 输入密码
#yanzhengma = input("请手动输入验证码：")
# ff.extract_image(url)
#yanzhengma = pyt3.image_file_to_string('code.png')[:4]

# 验证码处理

while b:
    driver.find_element_by_xpath(
        "//img[@id='jcaptcha']/..").click()  # 点击刷新验证码
    time.sleep(1)
    driver.save_screenshot('screenshot.png')
    imgelement = driver.find_element_by_id('jcaptcha')
    location = imgelement.location  # 获取验证码x,y轴坐标
    size = imgelement.size  # 获取验证码的长宽
    rangle = (int(location['x']), int(location['y']), int(
        location['x']+size['width']), int(location['y']+size['height']))  # 写成我们需要截取的位置坐标
    i = Image.open("screenshot.png")  # 打开截图
    result = i.crop(rangle)  # 使用Image的crop函数，从截图中再次截取我们需要的区域
    result.save('result.png')
    im=Image.open("result.png")
    imgry = im.convert('L')#图像加强，二值化
    sharpness =ImageEnhance.Contrast(imgry)#对比度增强
    sharp_img = sharpness.enhance(2.0)
    sharp_img.save("result.png")
    yanzhengma = pyt3.image_file_to_string('result.png', 'eng').replace(' ','')[:4]
    print("> %s <" % yanzhengma)
    break

    



driver.find_element_by_xpath(my.xpath3).clear() #清空验证码输入栏
driver.find_element_by_xpath(my.xpath3).send_keys(yanzhengma)  # 输入验证码
driver.find_element_by_xpath(my.xpath4).click()  # 点击登陆
time.sleep(2)

driver.find_element_by_xpath(my.xpath_zxt).click() #点击子系统
time.sleep(1)
driver.find_element_by_xpath(my.xpath_ymfw).click() #点击域名访问
    ## 登陆成功
    #try:
    #    driver.find_element_by_xpath(
    #        "//div[contains(text(),'验证码不正确,请重新输入验证码')]")
    #    time.sleep(1)
    #    driver.find_element_by_xpath(
    #        "//button[@class = 'swal2-confirm swal2-styled']").click()
    #except:
    #    b = False


# 加载时间过长 异常处理
time.sleep(2)
driver.find_element_by_xpath(my.xpath5).click()
time.sleep(1)
driver.find_element_by_xpath(my.xpath6).click()
time.sleep(1)
driver.find_element_by_xpath(my.xpath7).click()
time.sleep(1)
driver.find_element_by_xpath(my.xpath8).click()  # 点击贫困户维护页面
time.sleep(2)
WebDriverWait(driver, 10).until(EC.presence_of_element_located(
    (By.XPATH, "//div[div[@class = 'loader'] and starts-with(@style,'display: none')]")))

# 显示查询画面
# driver.find_element_by_xpath("//input[@formcontrolname='aab002']").send_keys(u'张梅')
# 输入查询人姓名
# time.sleep(2)
#！！出现重名，所以弃用

# p1 =
# personData.personData("李明伏","432424196404165015","13875184544","6217995580015235015","王大厦","13974204998","96")
# p1 =
# personData.personData("李明伏","432424196404165015","13875184544","6217995580015235015","王自平","13974204998","96")
# e_to_j.save_as_json(list_poor_family,'002.json')

    # for p1 in list(reversed(list_poor_family[:])):


#try:
for f in familynums:
    p1 = sqlff.load_sql_family(csor,*f)
    if (not(p1.state) and not(p1.error)):
        #if int(p1.ID)%20==0:
        #    e_to_j.personDatalist_to_json(list_poor_family, './file/fp_on.json')
        time.sleep(1)
        try:
            driver.find_element_by_xpath(my.xpath9).clear()  # 清空
            driver.find_element_by_xpath(
                my.xpath9).send_keys(p1.o5)  # 输入查询人身份证号
            time.sleep(1)
            driver.find_element_by_xpath(my.xpath10).click()  # 点击查询按钮
            time.sleep(2)
            #！！如果查询不到 异常处理 记录日志 保存日志
            # driver.find_element_by_xpath("//*[@id="ui-tabpanel-1"]/div/busi-tab/object-poor-family/p-panel[3]/div/div[2]/div/div/object-poor-family-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr/td[4]/span/span").click()
            # 点击第一栏查询项姓名栏（如果查询结果唯一）
            # time.sleep(2)
            # 添加显式等待
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//div[div[@class = 'loader'] and starts-with(@style,'display: none')]")))
            # driver.find_element_by_xpath(my.xpath11).click()
            # 点击第一栏查询项姓名栏（如果查询结果唯一）
    
            # 点击第一栏查询项姓名栏（如果查询结果唯一）
            ff.retryingFindClick(my.xpath11, driver)
            time.sleep(1)
    
        except (ElementNotVisibleException, NoSuchElementException, TimeoutException)as e:
            print("009")
            p1.add_log_e1(e)
            p1.show_error()
            error_count += 1
            continue
    
        #户基础信息修改开始
    
        ff.radio_driver_zp(driver,p1.a5,my.a5)#致贫原因1
        ff.radio_driver_zp(driver,p1.a6,my.a6)#致贫原因2
        ff.radio_driver_zp(driver,p1.a7,my.a7)#致贫原因3
        ff.input_driver(driver,p1.a1,my.a1)#联系电话
        ff.input_driver(driver,p1.a2,my.a2)#开户银行
        ff.input_driver(driver,p1.a3,my.a3)#银行卡号
        driver.find_element_by_xpath(my.xpath_qiehuan_3).click()  # 切换至三、
        time.sleep(2)
        ff.input_driver(driver,p1.a8,my.a8)#工资性收入
        ff.input_driver(driver,p1.a9,my.a9)#生产经营性收入
        #ff.input_driver(driver,p1.a10,my.a10)
        ff.input_driver(driver,p1.a11,my.a11)#资产收益分红型收入
        ff.input_driver(driver,p1.a12,my.a12)#其他财产性收入
        ff.input_driver(driver,p1.a13,my.a13)#计划生育金
        ff.input_driver(driver,p1.a14,my.a14)#低保金
        ff.input_driver(driver,p1.a15,my.a15)#特困供养金
        ff.input_driver(driver,p1.a16,my.a16)#生产经营性支出
        ff.input_driver(driver,p1.a17,my.a17)#养老保险金
        ff.input_driver(driver,p1.a18,my.a18)#生态补偿金
        ff.input_driver(driver,p1.a19,my.a19)#其他转移性收入
        driver.find_element_by_xpath(my.xpath_qiehuan_2).click()  # 切换至二、
        time.sleep(2)
        ff.input_driver(driver,p1.a20,my.a20)#耕地面积
        ff.input_driver(driver,p1.a21,my.a21)#林地面积
        ff.input_driver(driver,p1.a22,my.a22)#退耕还林面积
        ff.input_driver(driver,p1.a23,my.a23)#林果面积
        ff.input_driver(driver,p1.a24,my.a24)#牧草地面积
        ff.input_driver(driver,p1.a25,my.a25)#水面面积
        ff.input_driver(driver,p1.a29,my.a29)#与村主干路距离
        ff.input_driver(driver,p1.a31,my.a31)#住房面积
        ff.select_driver(driver,p1.a26,my.a26)#是否加入农村合作社
        ff.select_driver(driver,p1.a27,my.a27)#是否有龙头企业带动 
        ff.select_driver(driver,p1.a28,my.a28)#是否有创业致富带头人带动
        ff.select_driver(driver,p1.a30,my.a30)#入户路类型
        ff.select_driver(driver,p1.a32,my.a32)#是否通生活用电
        ff.select_driver(driver,p1.a33,my.a33)#是否通广播电视
        ff.select_driver(driver,p1.a34,my.a34)#是否解决安全饮用水
        #ff.select_driver(driver,p1.a37,my.a37)#危房户
        #ff.select_driver(driver,p1.a38,my.a38)#危房级别  
        p1.addover(csor)#标记状态


        driver.find_element_by_xpath(my.xpath_qiehuan_1).click()  # 切换至一、
        ps = sqlff.load_sql_person(csor,*f)

        for p2 in ps:
            ff.radio_driver(driver,p2.o4,my.xpath37)
            time.sleep(1.5)
            ff.select_driver(driver,p2.b1,my.b1)#与户主关系 
            ff.select_driver(driver,p2.b2,my.b2)#民族
            ff.select_driver(driver,p2.b3,my.b3)#政治面貌
            ff.select_driver(driver,p2.b5,my.b5)#文化程度
            ff.select_driver(driver,p2.b4,my.b4)#在校生状况
            ff.select_driver(driver,p2.b7,my.b7)#劳动技能
            #ff.select_driver(driver,p2.b9,my.b9)#省
            #ff.select_driver(driver,p2.b10,my.b10)#市
            #ff.select_driver(driver,p2.b11,my.b11)#县
            #ff.select_driver(driver,p2.b12,my.b12)#镇
            #ff.input_driver(driver,p2.b13,my.b13)#务工时间
            ff.select_driver(driver,p2.b14,my.b14)#失学或辍学原因
            ff.select_driver(driver,p2.b15,my.b15)#是否会讲普通话
            ff.select_driver(driver,p2.b16,my.b16)#是否享受农村居民最低生活保障
            ff.select_driver(driver,p2.b17,my.b17)#是否享受人身意外保险补贴
            ff.select_driver(driver,p2.b18,my.b18)#是否参加商业补充医疗保险
            ff.select_driver(driver,p2.b19,my.b19)#是否参加城乡居民基本养老保险
            ff.select_driver(driver,p2.b20,my.b20)#是否参加城乡居民基本医疗保险
            ff.select_driver(driver,p2.b21,my.b21)#是否参加大病保险
            p2.addover(csor)#标记状态

            cnn.commit()#提交数据库更改
    
        #人数据修改完成
        driver.find_element_by_xpath(
                    my.xpath27).click()  # 点击保存按钮
        time.sleep(1)                                
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            (By.XPATH, "//div[div[@class = 'loader'] and starts-with(@style,'display: none')]")))
        try:
            time.sleep(0.5)
            driver.execute_script(js_below)  # js点击上层按钮
            # driver.find_element_by_xpath(my.xpath26).click()#点击上层确定按钮
        except:
            print("001")
            try:
                driver.execute_script(js_below)  # js点击上层按钮
            except (ElementNotVisibleException, NoSuchElementException, TimeoutException) as e:
                print("001错误 %s" % e)
        time.sleep(0.5)
        driver.find_element_by_xpath(
                    my.xpath28).click()
        # 户基础信息修改完毕
        #p1.show_state()
        continue  #终止修改结对帮扶人信息
    
        # 开始修改信息
        try:
            driver.find_element_by_xpath(
                my.xpath14).click()  # 点击帮扶责任人结对信息栏
        except ElementNotVisibleException as e:
    
            print(e)
    
        time.sleep(2)
    
        try:  # 判定是否发现该结对帮扶人
            time.sleep(3)
            driver.find_element_by_xpath(
                "//span[contains(text(),'%s')]" % p1.helpPerson)
            a = True
        except:
            a = False
    
        if a:
            # 通过姓名定位后获取编号和电话号码
            # 如果联系电话一致，且开始日期不为"2018年06年01日"，则需要先判定序号，为1时直接执行修改日期程序，不为1时，先js注入，点击radio，再执行修改程序。保存后。保存edit
            if driver.find_element_by_xpath("//span[contains(text(),'%s')]/../../td[6]/span/span" % p1.helpPerson).text != p1.startdate and driver.find_element_by_xpath("//span[contains(text(),'%s')]/../../td[7]/span/span" % p1.helpPerson).text != p1.enddate:
                                                                                                                                                                                                # log
                                                                                                                                                                                                                                # driver.find_element_by_xpath("//span[contains(text(),'%s')]/../../td[14]/span"%
                                                                                                                                                                                                                                # p1.helpPerson).text == p1.helpPerson_phone and
                                                                                                                                                                                                                                # driver.find_element_by_xpath("//span[contains(text(),'%s')]/../../td[6]/span/span"%
                                                                                                                                                                                                                                # p1.helpPerson).text != p1.startdate
    
                # 判定序号
                if driver.find_element_by_xpath("//span[contains(text(),'%s')]/../../td[14]/span" % p1.helpPerson).text != "1":
                    # 注入js 模拟点击radio
                    # js =
                    #"$('object-poor-family-twinning-grid>p-datatable>div>div>div>:nth-child(2)>div>table>:nth-child(2)>:nth-child(%d)>td>p-dtradiobutton>div>div:nth-child(2)>span').click()"
                    #%
                    # int(driver.find_element_by_xpath("//span[contains(text(),'%s')]/../../td[14]/span"
                    #% p1.helpPerson).text)
                    # driver.execute_script(js)
    
                    # 也可以直接点击input下的span
                    try:
                        driver.find_element_by_xpath(
                            "//span[contains(text(),'%s')]/../../td[1]/*/*/div[2]/span" % p1.helpPerson).click()
                    except ElementNotVisibleException as e:
                        pass
    
                driver.find_element_by_xpath(
                    my.xpath17).click()  # 点击修改日期按钮
                time.sleep(1.5)
                driver.find_element_by_xpath(
                    my.xpath18).clear()  # 清除开始日期
                driver.find_element_by_xpath(
                    my.xpath18).send_keys(p1.startdate)  # 输入开始日期
                driver.find_element_by_xpath(
                    my.xpath19).clear()  # 清除结束日期
                driver.find_element_by_xpath(
                    my.xpath19).send_keys(p1.enddate)  # 输入结束日期
                driver.find_element_by_xpath(
                    my.xpath20).click()  # 点击保存按钮
                time.sleep(1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, "//div[div[@class = 'loader'] and starts-with(@style,'display: none')]")))
                try:
                    time.sleep(0.5)
                    driver.execute_script(js_below)  # js点击上层按钮
                    # driver.find_element_by_xpath(my.xpath26).click()#点击上层确定按钮
                except:
                    print("001")
                    try:
                        driver.execute_script(js_below)  # js点击上层按钮
                    except (ElementNotVisibleException, NoSuchElementException, TimeoutException) as e:
                        print("001错误 %s" % e)
                    # driver.find_element_by_xpath(my.xpath36).click()#点击
                    # 添加修改帮扶日期页面关闭X按钮
                    # time.sleep(0.5)
                    driver.find_element_by_xpath(
                        my.xpath28).click()  # 点击贫困户信息页关闭按钮
                    time.sleep(0.5)
    
                    p1.show_editdate()
                    continue
    
                time.sleep(0.5)
                driver.find_element_by_xpath(
                    my.xpath28).click()  # 点击贫困户信息页关闭按钮
    
                p1.show_editdate()
                continue
            else:  # 无需修改
                if p1.error == True:
                    print("该户上次录入时出错，跳过，待手工处理（）")
                driver.find_element_by_xpath(
                    my.xpath28).click()  # 点击贫困户信息页关闭按钮
                p1.pass_state()
                continue
    
        else:  # 否则点击新增结对帮扶人
            try:
                driver.find_element_by_xpath(
                    my.xpath16).click()  # 新增结对
                time.sleep(1.5)
                driver.find_element_by_xpath(my.xpath21).send_keys(
                    p1.helpPerson)  # 输入结对帮扶人姓名
                driver.find_element_by_xpath(
                    my.xpath22).click()  # 点击查询
                WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, "//div[div[@class = 'loader'] and starts-with(@style,'display: none')]")))
                time.sleep(0.5)
                driver.find_element_by_xpath(
                    my.xpath24).send_keys(p1.startdate)
                driver.find_element_by_xpath(
                    my.xpath25).send_keys(p1.enddate)
                try:
                    driver.find_element_by_xpath(
                        my.xpath23).click()  # 点击保存
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, "//div[div[@class = 'loader'] and starts-with(@style,'display: none')]")))
                    time.sleep(1)
                    try:
                        driver.execute_script(js_below)  # js点击上层按钮
                        time.sleep(0.5)
                    except:
                        print("0031")
                        p1.add_log_e0()
                        p1.show_error()
                        try:
                            driver.find_element_by_xpath(
                                my.xpath34).click()  # 添加结对帮扶责任人列表关闭X按钮
                            time.sleep(1)
                            driver.find_element_by_xpath(
                                my.xpath28).click()  # 点击贫困户信息页关闭按钮
                            continue
                        except:
                            print("........")
    
                    time.sleep(1)
                    driver.find_element_by_xpath(
                        my.xpath28).click()  # 点击贫困户信息页关闭按钮
                    time.sleep(0.5)
                    p1.show_edit()
                    continue
    
                    # try:
                    #     driver.execute_script(js_below) # js点击上层按钮
                    # except :
                    #     pass
    
                except (ElementNotVisibleException, NoSuchElementException, TimeoutException)as e:
                    print("003")
                    p1.add_log_e0()
                    p1.show_error()
    
                    try:
                        driver.find_element_by_xpath(
                            my.xpath34).click()  # 添加结对帮扶责任人列表关闭X按钮
                    except:
                        print("........")
    
                    time.sleep(1)
                    driver.find_element_by_xpath(
                        my.xpath28).click()  # 点击贫困户信息页关闭按钮
                    time.sleep(0.5)
                    continue  # c
    
            except(ElementNotVisibleException, NoSuchElementException, TimeoutException)as e:
                print("004")
                p1.add_log_e2()
                p1.show_error()
                continue
    
        try:
            # driver.find_element_by_xpath(my.xpath27).click()#点击页面保存
            # time.sleep(1)
            # try:
            #    driver.find_element_by_xpath(my.xpath26).click()#点击确定
            # except :
            #    pass
            time.sleep(0.5)
            driver.find_element_by_xpath(
                my.xpath28).click()  # 点击贫困户信息页关闭按钮
            time.sleep(0.5)
            continue
    
        except:
    
            continue
    else:
        p1.pass_state()
#except(ElementNotVisibleException, NoSuchElementException, TimeoutException)as e:
#    print(e)
#finally:
#    e_to_j.personDatalist_to_json(list_poor_family, './file/fp_on.json')
    
#    if input("程序结束，是否导出错误列表Y/N").lower()=="y":
#        e_to_j.js_to_xlsx("./file/fp_on.json","./file/error.xlsx")




    # 移除按钮的隐藏类名，使其可见
    # js =
    #"document.querySelector('object-poor-family-twinning-grid>p-datatable>div>div>div>:nth-child(2)>div>table>:nth-child(2)>:nth-child(%d)>td>p-dtradiobutton>div>div:nth-child(1)').removeClass"%
    # n N表示第几行
    # driver.execute_script(js)

    #:nth-child(n) 第N个子元素

    # object-poor-family-twinning-grid>p-datatable>div>div>div>:nth-child(2)>div>table>:nth-child(2)>:nth-child(1)>td>p-dtradiobutton>div>div
    # 已结对帮扶责任人列表第一个radioselector 用于移除class
    # object-poor-family-twinning-grid>p-datatable>div>div>div>:nth-child(2)>div>table>:nth-child(2)>:nth-child(2)>td>p-dtradiobutton>div>div
    # 已结对帮扶责任人列表第一个radioselector 用于移除class
    # object-poor-family-twinning-grid>p-datatable>div>div>div>:nth-child(2)>div>table>:nth-child(2)>:nth-child(3)>td>p-dtradiobutton>div>div
    # 已结对帮扶责任人列表第一个radioselector 用于移除class

    # //object-poor-family-twinning-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr[%d]/td[2]/span/span
    # % n 第n栏编号元素的xpath(编号1,2,3)
    # //object-poor-family-twinning-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr[%d]/td[3]/span
    # % n 第n栏 姓名 元素的xpath
    # //object-poor-family-twinning-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr[%d]/td[4]/span
    # % n 第n栏 单位 元素的xpath
    # //object-poor-family-twinning-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr[%d]/td[6]/span/span
    # % n 第n栏 开始日期 元素的xpath
    # //object-poor-family-twinning-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr[%d]/td[7]/span/span
    # % n 第n栏 结束日期 元素的xpath
    # //object-poor-family-twinning-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr[%d]/td[14]/span
    # % n 第n栏 联系电话 元素的xpath
    # driver.find_elements_by_xpath("").text 获取元素的文本值
    # //span[contains(text(),'%s')]/../../td[2]/span/span 根据姓名查找序号
    # //span[contains(text(),'%s')]/../../td[14]/span 根据姓名查找联系电话
    # //span[contains(text(),'%s')]/../../td[6]/span/span 根据姓名查找开始结对日期
    # //span[contains(text(),'孙军')]/../../td[1]/*/*/*/* 根据姓名查找radio/input
    # 相邻元素定位，
    # 前一位：
    # preceding-sibling::div[1]
    # 后一位：
    # following-sibling::div[1]
    # 前N位：
    # preceding-sibling::div[N]
    # 后N位：
    # following-sibling::div[N]

    # 思路1：破解display:none问题
    # 先通过xpath文本查找（帮扶人姓名、电话、开始、结束他日期）确定需要修改的行号，然后根据行号使用js（CSSselector）注入移除类名，实现点击radio

    # 思路2：通过xpath文本查找（帮扶人姓名、电话、开始、结束他日期）判定是否需要修改，同时确定行号后，删除多余行，使待编辑行变为首行"""


# driver.find_element_by_xpath("//*[@id="ui-tabpanel-1"]/div/busi-tab/object-poor-family/p-panel[3]/div/div[2]/div/div/object-poor-family-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr[1]/td[4]/span/span/a").click()#点击第一栏查询项姓名栏（如果查询结果不唯一）
# time.sleep(2)

# xpath（//span[contains(text(),'五、帮扶责任人结对信息')]）


#  //button[@label='取消结对']/span
#  //button[@label='增加结对']/span
#  //button[@label='修改时间']/span

# 修改日期 1.开始日期 label //input[@class='ng-tns-c6-73 ui-inputtext ui-widget
# ui-state-default ui-corner-all']
#          2.结束日期 label //input[@class='ng-tns-c6-74 ui-inputtext ui-widget
#          ui-state-default ui-corner-all']
#          3.保存修改日期 button //button[@id = 'saveTime' and @label='保存']

# 新增结对帮扶人姓名 1.输入结对帮扶人姓名 //input[@formcontrolname='aab002']
#                   2.查询 button //button[@id = 'queryByName' and @label='查询']
#                   3.开始日期 //input[@class='ng-tns-c6-71 ui-inputtext ui-widget
#                   ui-state-default ui-corner-all']
#                   4.结束日期 //input[@class='ng-tns-c6-72 ui-inputtext ui-widget
#                   ui-state-default ui-corner-all']
#  如果查询无结果 异常处理


#  整体确定 1.button //button[@class='swal2-confirm swal2-styled']

###


# excel转json

# 循环体


# e_to_j.excel_json(r"C:\Users\wangz\Desktop\东岳村
# 贫困户信息_20180814.xlsx",r"C:\Users\wangz\Desktop\2.json") #转换
    #e_to_j.personDatalist_to_json(list_poor_family, '002.json')

    #with open("log.txt", 'w', encoding="utf-8") as f:
    #    for p1 in list_poor_family:
    #        f.writelines(p1.log+'\n')
    

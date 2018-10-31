# -*- coding:utf-8 -*-
import json
import codecs
class setting():
    """w"""
    def __init__(self,account,password):
        self.url = "http://cpadisc4.cpad.gov.cn/cpad/login"
        self.account = account
        self.password = password
        self.chromePath = u'C:/Users/wangz/AppData/Local/Google/Chrome/Application/chromedriver'
        self.chromePath_old = u'C:/Program Files (x86)/Google/Chrome/Application/chromedriver'
        self.xpath1 = "//input[@name='username']"#帐号输入框体
        self.xpath2 = "//input[@name='password']"#密码输入框体
        self.xpath3 = "//input[@name='jcaptcha_response']"#验证码输入框体
        self.xpath4 = "//div[@id='loginBut']"#首页登陆按钮

        self.xpath_zxt = "//a[contains(text(),'业务管理子系统')]"
        self.xpath_ymfw = "//a[contains(text(),'域名访问')]"
        self.xpath5 = "//a[@title='扶贫对象']"#扶贫对象菜单按钮
        self.xpath6 = "//a[@title='基础信息维护']"#扶贫对象=>基础信息子菜单按钮
        self.xpath7 = "//a[@title='2018年度']"#扶贫对象=>基础信息=>2018年度子菜单按钮
        self.xpath8 = "/html/body/app-root/div/div/app-main-layout/div/div[1]/div/div/app-menu/ul/nui-main-menu/div/div[1]/div[2]/div/nui-main-menu-sub/ul/li/nui-main-menu-sub/ul/li[6]/nui-main-menu-sub/ul/li[1]/a/span[2]"#扶贫对象=>基础信息=>2018年度=>贫困户按钮
        self.xpath9 = "//input[@formcontrolname='aab004']"#基础信息维护界面=>身份证号查询输入框体
        self.xpath10 = "//button[@label='查询']/span"#基础信息维护界面=>查询按钮
        self.xpath11 = "//p-datatable[@datakey]/*/*/*/div[2]/*/*/tbody/*/td[4]/*/span"#查询结果第一栏
        self.xpath12 = "//input[@formcontrolname='aar012']"#贫困户首页基础信息=>联系电话框体
        self.xpath13 = "//input[@formcontrolname='aac004']"#贫困户首页基础信息=>银行卡号框体
        self.xpath14 = "//span[contains(text(),'五、帮扶责任人结对信息')]"#结对帮扶人信息选项卡选择
        self.xpath15 = "//button[@label='取消结对']/span"#取消结对按钮
        self.xpath16 = "//button[@label='增加结对']/span"#增加结对按钮
        self.xpath17 = "//button[@label='修改时间']/span"#修改时间按钮
        self.xpath18 = "//object-poor-family/p-dialog[4]/div/div[2]/p-tabview/div/div/form/div/div/div[2]/nui-date-nav/div/p-calendar/span/input"#修改日期=>开始日期框体
        self.xpath19 = "//object-poor-family/p-dialog[4]/div/div[2]/p-tabview/div/div/form/div/div/div[4]/nui-date-nav/div/p-calendar/span/input"#修改日期=>结束日期框体
        self.xpath20 = "//button[@id = 'saveTime' and @label='保存']"#修改日期=>保存按钮
        self.xpath21 = "//input[@placeholder='请输入帮扶责任人姓名']"#新增结对帮扶人=>输入结对帮扶人姓名框体
        self.xpath22 = "//button[@id = 'queryByName' and @label='查询']"#新增结对帮扶人=>查询按钮
        self.xpath23 = "//button[@id='saveTwinning']/span"#新增结对帮扶人=>保存按钮
        self.xpath24 = "//object-poor-family-twinning-addgrid/form[2]/div/div/div[2]/nui-date-nav/div/p-calendar/span/input"#新增结对帮扶人=>开始日期框体
        self.xpath25 = "//object-poor-family-twinning-addgrid/form[2]/div/div/div[4]/nui-date-nav/div/p-calendar/span/input"#新增结对帮扶人=>结束日期框体
        self.xpath26 = "//button[@class='swal2-confirm swal2-styled']"#上层确定按钮
        self.xpath27 = "//button[@id='on_save']/span"#页面保存按钮
        self.xpath28 = "//button[@id='on_cancel']/span"#页面关闭按钮
        self.xpath29 = "//object-poor-family-twinning-addgrid/p-datatable/div/div/div/div[2]/div/table/tbody/tr/td/p-dtradiobutton/div/div/input"#新增结对帮扶人选择框1号
        self.xpath30 = "//object-poor-family-twinning-addgrid/p-datatable/div/div/div/div[2]/div/table/tbody/tr[2]/td/p-dtradiobutton/div/div/input"#新增结对帮扶人选择框2号
        self.xpath31 = "//object-poor-family-twinning-addgrid/p-datatable/div/div/div/div[2]/div/table/tbody/tr/td[3]/span"#新增结对帮扶人姓名（.text）
        self.xpath32 = "//object-poor-family-twinning-addgrid/p-datatable/div/div/div/div[2]/div/table/tbody/tr/td[14]/span"#新增结对帮扶人联系电话（.text）
        self.xpath33 = "//object-poor-family-twinning-grid/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr/td[3]/span"
        self.xpath34 = "//span[contains(text(),'添加帮扶责任人列表')]/../a"#添加结对帮扶责任人列表关闭X按钮
        self.xpath35 = "//span[contains(text(),'贫困户信息')]/../a"#贫困户信息关闭X按钮
        self.xpath36 = "//span[contains(text(),'修改帮扶时间')]/../a" #添加修改帮扶日期页面关闭X按钮

        self.xpath_qiehuan_1="//span[contains(text(),'一、基本情况')]"#一、基本情况选项卡选择
        self.xpath_qiehuan_2="//span[contains(text(),'二、生产生活条件')]"#二、生产生活条件选项卡选择
        self.xpath_qiehuan_3="//span[contains(text(),'三、上年度收入、低保和患病信息')]"#三、上年度收入、低保和患病信息选项卡选择
        

        self.xpath37 = "//p-panel[@header ='家庭成员']/descendant::*//span[contains(text(),'?')]/../../../td[1]/*/*/div[2]/span"  #根据家庭成员姓名获取前排的radio点击单元
        self.a1="//input[@formcontrolname='aar012']"#贫困户首页基础信息=>联系电话框体
        self.a2 ="//input[@formcontrolname='aaq002']" #开户银行
        self.a3 = "//input[@formcontrolname='aac004']"#贫困户首页基础信息=>银行卡号框体
        self.a4=""#计划脱贫年度
        self.a5="//radio-checkbox-code[@groupname='aac007']/*/*/div[?]/*/*/div[2]/span"
        #致贫原因1 1因病 2因残 3因学 4因灾 5缺土地 6缺水 7缺技术 8缺劳力 9缺资金 10交通条件落后 11自身发展动力不足 12因婚 13因丧
        self.a6="//radio-checkbox-code[@groupname='aac008_2']/*/*/div[?]/*/*/div[2]/span"#致贫原因2
        self.a7="//radio-checkbox-code[@groupname='aac008_3']/*/*/div[?]/*/*/div[2]/span"#致贫原因3
        self.a7_1 =""#返贫原因
        self.a8="//input[@formcontrolname='aac073']"#工资性收入
        self.a9="//input[@formcontrolname='aac071']"#生产经营性收入
        self.a10="//input[@formcontrolname='aac072']"#财产性收入
        self.a11="//input[@formcontrolname='aac092']"#资产收益分红型收入
        self.a12="//input[@formcontrolname='aac093']"#其他财产性收入
        self.a13="//input[@formcontrolname='aac076']"#计划生育金
        self.a14="//input[@formcontrolname='aac077']"#低保金
        self.a15="//input[@formcontrolname='aac086']"#特困供养金
        self.a16="//input[@formcontrolname='aac074']"#生产经营性支出
        self.a17="//input[@formcontrolname='aac087']"#养老保险金
        self.a18="//input[@formcontrolname='aac078']"#生态补偿金
        self.a19="//input[@formcontrolname='aac083']"#其他转移性收入
        self.a20="//input[@formcontrolname='aac301']"#耕地面积
        self.a21="//input[@formcontrolname='aac303']"#林地面积
        self.a22="//input[@formcontrolname='aac304']"#退耕还林面积
        self.a23="//input[@formcontrolname='aac305']"#林果面积
        self.a24="//input[@formcontrolname='aac306']"#牧草地面积
        self.a25="//input[@formcontrolname='aac307']"#水面面积
        self.a26="//p-dropdown[@id='aac084']/*/div[2]/span"#是否加入农村合作社 下拉框    //p-dropdown[@id='aac084']@class 点选 //span[@class='待获取'][contains(text(),'否')]  left12
        self.a27="//p-dropdown[@id='aac088']/*/div[2]/span"#是否有龙头企业带动    点选 //span[@class='ng-tns-c5-62'][contains(text(),'否')]
        self.a28="//p-dropdown[@id='aac089']/*/div[2]/span"#是否有创业致富带头人带动  //span[@class='ng-tns-c5-63'][contains(text(),'否')]
        self.a29="//input[@formcontrolname='aac315']"#与村主干路距离
        self.a30="//p-dropdown[@id='aac316']/*/div[2]/span"#入户路类型   //span[@class='ng-tns-c5-64'][contains(text(),'？？')]  泥土路 砂石路 水泥路 沥青路
        self.a31="//input[@formcontrolname='aac317']"#住房面积  
        self.a32="//p-dropdown[@id='aac313']/*/div[2]/span"#是否通生活用电  //span[@class='ng-tns-c5-65'][contains(text(),'是')] 
        self.a33="//p-dropdown[@id='aac314']/*/div[2]/span"#是否通广播电视  //span[@class='ng-tns-c5-66'][contains(text(),'是')] 
        self.a34="//p-dropdown[@id='aac091']/*/div[2]/span"#是否解决安全饮用水  //span[@class='ng-tns-c5-67'][contains(text(),'是')] 
        self.a35="//input[@formcontrolname='aac073']"#危房户
        self.a36="//input[@formcontrolname='aac073']"#危房级别
        self.a37="//p-dropdown[@id='aac320']/*/div[2]/span"#主要燃料类型  //span[@class='ng-tns-c5-69'][contains(text(),'是')] "柴草 干畜粪 煤炭 清洁能源 其它"
        self.a38="//p-dropdown[@id='aac319']/*/div[2]/span"#是否有卫生厕所   //span[@class='ng-tns-c5-70'][contains(text(),'是')] 
        self.b1="//p-dropdown[@id='aab006']/*/div[2]/span"#与户主关系 
        self.b2="//p-dropdown[@id='aab007']/*/div[2]/span"#民族
        self.b3="//p-dropdown[@id='aak033']/*/div[2]/span"#政治面貌
        self.b4="//p-dropdown[@id='aab008']/*/div[2]/span"#文化程度
        self.b5="//p-dropdown[@id='aab009']/*/div[2]/span"#在校生状况
        self.b6="//p-multiselect[@id='aab017']/*/div[3]/span"#健康状况   ##
        #b6+"/../../div[2]" .get_attribute("title").split(',') 得到已选的元素； b6+"/../../div[4]/div[2]//label[text()='健康']"
        self.b7="//p-dropdown[@id='aab010']/*/div[2]/span"#劳动技能
        self.b8="//p-dropdown[@id='aab006']/*/div[2]/span"#务工区域  xx
        self.b9="//p-dropdown[@id='aab025']/*/div[2]/span"#省
        self.b10="//p-dropdown[@id='aab026']/*/div[2]/span"#市
        self.b11="//p-dropdown[@id='aab027']/*/div[2]/span"#县
        self.b12="//p-dropdown[@id='aab028']/*/div[2]/span"#镇
        self.b13="//input[@formcontrolname='aab012']"#务工时间
        self.b14="//p-dropdown[@id='aab064']/*/div[2]/span"#失学或辍学原因
        self.b15="//p-dropdown[@id='aab065']/*/div[2]/span"#是否会讲普通话
        self.b16="//p-dropdown[@id='aab030']/*/div[2]/span"#是否享受农村居民最低生活保障
        self.b17="//p-dropdown[@id='aab066']/*/div[2]/span"#是否享受人身意外保险补贴
        self.b18="//p-dropdown[@id='aab067']/*/div[2]/span"#是否参加商业补充医疗保险
        self.b19="//p-dropdown[@id='aab014']/*/div[2]/span"#是否参加城乡居民基本养老保险
        self.b20="//p-dropdown[@id='aab013']/*/div[2]/span"#是否参加城乡居民基本医疗保险
        self.b21="//p-dropdown[@id='aab022']/*/div[2]/span"#是否参加大病保险
        

        

    
    def load_setting(self):
        with open('setting.json','r',encoding = 'utf-8') as f:
            list_f = json.load(f)
            self.account = list_f['account']
            self.password = list_f['password']
        print("已从'setting.json'加载帐号密码")

    def get_and_save_setting(self):
        account = input("请输入帐号")
        password = input("请输入密码")
        dict_j = {'account':account,'password':password}
        j_file = str(json.dumps(dict_j, ensure_ascii=False))
        with codecs.open('setting.json','w', encoding='utf-8', errors ='ignore') as f:
            f.write(j_file)
        print("已将账号和密码，保存至目录下的setting.json，可用记事本打开自行编辑，下次运行将自动读取已存配置")
            
from datetime import datetime
class personData():
    """储存待录入结对帮扶资料对象"""
    def __init__(self,ID,o1,o2,o3,o4,o5,o6,b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13,b14,b15,b16,b17,b18,b19,b20,b21,error = False,state = False,log=""):
        self.ID=ID
        self.o1=o1
        self.o2=o2
        self.o3=o3
        self.o4=o4
        self.o5=o5
        self.o6=o6
        self.b1=b1
        self.b2=b2
        self.b3=b3
        self.b4=b4
        self.b5=b5
        self.b6=b6
        self.b7=b7
        self.b8=b8
        self.b9=b9
        self.b10=b10
        self.b11=b11
        self.b12=b12
        self.b13=b13
        self.b14=b14
        self.b15=b15
        self.b16=b16
        self.b17=b17
        self.b18=b18
        self.b19=b19
        self.b20=b20
        self.b21=b21
        self.error=error
        self.state=state
        self.log=log


    def addover(self,csor):
        self.state = True
        sql = "UPDATE 'person' set state = 1 where o5 ='%s'" % self.o5
        csor.execute(sql)
        self.log = "完成，%s,%s,%s,%s人数据已录入系统。" % (datetime.now(),self.ID,self.o4,self.o5)
        sql2 = "UPDATE 'person' set log = '%s' where o5 ='%s'" % (self.log,self.o5)
        csor.execute(sql2)
        self.print_log()
        
    def adderror(self,csor):
        self.error = True
        sql = "UPDATE 'person' set error = 1 where o5 ='%s'" % self.o5
        csor.execute(sql)
        self.log = "错误，%s,%s,%s,%s人数据录入出错。" % (datetime.now(),self.ID,self.o4,self.o5)
        sql2 = "UPDATE 'person' set log = '%s' where o5 ='%s'" % (self.log,self.o5)
        csor.execute(sql2)
        self.print_log()

        


    def show_error(self):
        
        self.error = True
        self.print_log()
    
    def show_edit(self):
        self.state = True
        self.add_log_finish()
        self.print_log()

    def show_editdate(self):
        self.edit = True
        self.add_log_editdate()
        self.print_log()

    def end(self):
        self.state = True
        
        self.print_log()

    def add_log_e(self,e):
        self.log = "错误，%s,%s,%s,%s 出错，（%s）"% (datetime.now(),self.suoyin,self.name,self.ID,e)

    def add_log_e1(self,e):
        self.log = "错误，%s,%s,%s,%s 身份证号有误，无法进入基础信息修改界面，（%s）。"% (datetime.now(),self.suoyin,self.name,self.ID,e)

    def add_log_e2(self):
        self.log = "错误，%s,%s,%s,%s在新增结对帮扶责任人时出错。"% (datetime.now(),self.suoyin,self.name,self.ID)
    
    def add_log_e0(self):
        self.log = "错误，%s,%s,%s,%s，尚未查询到该结对帮扶责任人 %s 时出错。"% (datetime.now(),self.suoyin,self.name,self.ID,self.helpPerson)

    def add_log_same(self):
        self.log = "提示，%s,%s,%s,%s在系统中出现相同的结对帮扶人%s,标记为已处理，待手动处理。"% (datetime.now(),self.suoyin,self.name,self.ID,self.helpPerson)

    def add_log_finish(self):

        self.log = "完成，%s,%s,%s,%s 录入系统。"% (datetime.now(),self.ID,self.o4,self.o5)
    
    def add_log_editdate(self):

        self.log = "完成，%s,%s,%s,%s已将结对帮扶人 %s 时间修改完成。"% (datetime.now(),self.suoyin,self.name,self.ID,self.helpPerson)
    def add_pass_state(self):
        self.log = "提示：%s,%s,%s,%s状态为已修改，跳过..."% (datetime.now(),self.suoyin,self.name,self.ID)

    def print_log(self):
        print(self.log)

    def pass_state(self):
        self.edit = True
        self.add_pass_state()
        self.print_log()



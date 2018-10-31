from datetime import datetime
class familyData():
    """储存待录入结对帮扶资料对象"""
    def __init__(self,ID,o1,o2,o3,o4,o5,o6,o7,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,ps,state=False,error=False,log=""):
        self.ID = ID
        self.o1 = o1
        self.o2 = o2
        self.o3 = o3
        self.o4 = o4
        self.o5 = o5
        self.o6 = o6
        self.o7 = o7
        self.a1 = a1
        self.a2 = a2
        self.a3 = a3
        self.a4 = a4
        self.a5 = a5
        self.a6 = a6
        self.a7 = a7
        self.a8 = a8
        self.a9 = a9
        self.a10 = a10
        self.a11 = a11
        self.a12 = a12
        self.a13 = a13
        self.a14 = a14
        self.a15 = a15
        self.a16 = a16
        self.a17 = a17
        self.a18 = a18
        self.a19 = a19
        self.a20 = a20
        self.a21 = a21
        self.a22 = a22
        self.a23 = a23
        self.a24 = a24
        self.a25 = a25
        self.a26 = a26
        self.a27 = a27
        self.a28 = a28
        self.a29 = a29
        self.a30 = a30
        self.a31 = a31
        self.a32 = a32
        self.a33 = a33
        self.a34 = a34
        self.a35 = a35
        self.a36 = a36
        self.a37 = a37
        self.a38 = a38
        self.error = error
        self.log = log
        self.state = state
        self.ps =ps
        
    def addover(self,csor):
        self.state = True
        sql = "UPDATE 'family' set state = 1 where o3 ='%s'" % self.o3
        csor.execute(sql)
        self.log = "完成，%s,%s,%s,%s户数据已录入系统。" % (datetime.now(),self.ID,self.o4,self.o5)
        sql2 = "UPDATE 'family' set log = '%s' where o3 ='%s'" % (self.log,self.o3)
        csor.execute(sql2)
        self.print_log()
        
    def adderror(self,csor):
        self.error = True
        sql = "UPDATE 'family' set error = 1 where o3 ='%s'" % self.o3
        csor.execute(sql)
        self.log = "错误，%s,%s,%s,%s户数据录入出错。" % (datetime.now(),self.ID,self.o4,self.o5)
        sql2 = "UPDATE 'family' set log = '%s' where o3 ='%s'" % (self.log,self.o3)
        csor.execute(sql2)
        self.print_log()


    def show_error(self):
        
        self.error = True
        self.print_log()
    
    def show_state(self):
        self.state = True
        self.add_log_finish()
        self.print_log()

    def show_statedate(self):
        self.state = True
        self.add_log_statedate()
        self.print_log()

    def end(self):
        self.state = True
        
        self.print_log()

    def add_log_e(self,e):
        self.log = "错误，%s,%s,%s,%s 出错，（%s）" % (datetime.now(),self.ID,self.o4,self.ID,e)

    def add_log_e1(self,e):
        self.log = "错误，%s,%s,%s,%s 身份证号有误，无法进入基础信息修改界面，（%s）。" % (datetime.now(),self.ID,self.o4,self.ID,e)

    def add_log_e2(self):
        self.log = "错误，%s,%s,%s,%s在新增结对帮扶责任人时出错。" % (datetime.now(),self.ID,self.o4,self.o3)
    
    def add_log_e0(self):
        self.log = "错误，%s,%s,%s,%s，尚未查询到该结对帮扶责任人 %s 时出错。" % (datetime.now(),self.ID,self.o4,self.ID,self.helpPerson)

    def add_log_same(self):
        self.log = "提示，%s,%s,%s,%s在系统中出现相同的结对帮扶人%s,标记为已处理，待手动处理。" % (datetime.now(),self.ID,self.o4,self.ID,self.helpPerson)

    def add_log_finish(self):

        self.log = "完成，%s,%s,%s,%s已录入系统。" % (datetime.now(),self.ID,self.o4,self.ID)
    
    def add_log_statedate(self):

        self.log = "完成，%s,%s,%s,%s已将结对帮扶人 %s 时间修改完成。" % (datetime.now(),self.ID,self.o4,self.ID,self.helpPerson)
    def add_pass_state(self):
        self.log = "提示：%s,%s,%s,%s状态为已修改，跳过..." % (datetime.now(),self.ID,self.o4,self.ID)

    def print_log(self):
        print(self.log)

    def pass_state(self):
        self.state = True
        self.add_pass_state()
        self.print_log()





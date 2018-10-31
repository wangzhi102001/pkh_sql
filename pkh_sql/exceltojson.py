
import xlrd
from collections import OrderedDict
import json
import codecs
import personData
import familyData
import datetime
import xlsxwriter

 



def save_as_json(object,file_path): 
    '''将数据存储到json文件'''  #第三步
    with open(file_path,"w",encoding='utf-8',errors='ignore') as f:
        json.dump(object,f,ensure_ascii=False)


def file_to_json_fomat(path1,path2,path3,list_poor_family):  #第二步
    '''将json文件转化成标准有序的字典'''
    with open(path1,'r',encoding ="utf-8")as f:#加载json文件
        familyDatas =json.load(f)#将加载的json文件转换成字典列表
    with open(path2,'r',encoding ="utf-8")as p:#加载json文件
        personDatas =json.load(p)#将加载的json文件转换成字典列表

    for family in familyDatas:
        list_p=[]
        for person in personDatas:
            if person['o3']==family['o3']:
                list_p.append({'ID':person['ID'],'o1':person['o1'],'o2':person['o2'],'o3':person['o3'],'o4':person['o4'],'o5':person['o5'],'o6':person['o6'],'b1':person['b1'],'b2':person['b2'],'b3':person['b3'],'b4':person['b4'],'b5':person['b5'],'b6':person['b6'],'b7':person['b7'],'b8':person['b8'],'b9':person['b9'],'b10':person['b10'],'b11':person['b11'],'b12':person['b12'],'b13':person['b13'],'b14':person['b14'],'b15':person['b15'],'b16':person['b16'],'b17':person['b17'],'b18':person['b18'],'b19':person['b19'],'b20':person['b20'],'b21':person['b21'],'error':False,'state':False,'log':person['log']})
        list_poor_family.append({'ID':family['ID'],'o1':family['o1'],'o2':family['o2'],'o3':family['o3'],'o4':family['o4'],'o5':family['o5'],'o6':family['o6'],'o7':family['o7'],'a1':family['a1'],'a2':family['a2'],'a3':family['a3'],'a4':family['a4'],'a5':family['a5'],'a6':family['a6'],'a7':family['a7'],'a8':family['a8'],'a9':family['a9'],'a10':family['a10'],'a11':family['a11'],'a12':family['a12'],'a13':family['a13'],'a14':family['a14'],'a15':family['a15'],'a16':family['a16'],'a17':family['a17'],'a18':family['a18'],'a19':family['a19'],'a20':family['a20'],'a21':family['a21'],'a22':family['a22'],'a23':family['a23'],'a24':family['a24'],'a25':family['a25'],'a26':family['a26'],'a27':family['a27'],'a28':family['a28'],'a29':family['a29'],'a30':family['a30'],'a31':family['a31'],'a32':family['a32'],'a33':family['a33'],'a34':family['a34'],'a35':family['a35'],'a36':family['a36'],'a37':family['a37'],'a38':family['a38'],'ps':list_p,'error':False,'state':False,'log':family['log']})

    save_as_json(list_poor_family,path3)



def json_to_familyDatalist(path,list_js,lista,error,start,end,n):  #初始化数据
    '''读取之前或最新的json，并加载成可操作的family对象列表'''
    with open(path,'r',encoding ="utf-8")as f:#加载json文件
        list_js =json.load(f)#将加载的json文件转换成字典列表
#input()
    
    errorp=0
    endp=0
    startp=0
    np=0
    for family in list_js:#将加载的json数据添加到familyData列表中
        list_p =[]
        for person in family['ps']:            
            list_p.append(personData.personData(person['ID'],person['o1'],person['o2'],person['o3'],person['o4'],person['o5'],person['o6'],person['b1'],person['b2'],person['b3'],person['b4'],person['b5'],person['b6'],person['b7'],person['b8'],person['b9'],person['b10'],person['b11'],person['b12'],person['b13'],person['b14'],person['b15'],person['b16'],person['b17'],person['b18'],person['b19'],person['b20'],person['b21'],person['error'],person['state'],person['log']))
            if not(person['state']):
                if person['error']:
                    errorp+=1
                startp+=1
            else:
                endp+=1
            np+=1

        lista.append(familyData.familyData(family['ID'],family['o1'],family['o2'],family['o3'],family['o4'],family['o5'],family['o6'],family['o7'],family['a1'],family['a2'],family['a3'],family['a4'],family['a5'],family['a6'],family['a7'],family['a8'],family['a9'],family['a10'],family['a11'],family['a12'],family['a13'],family['a14'],family['a15'],family['a16'],family['a17'],family['a18'],family['a19'],family['a20'],family['a21'],family['a22'],family['a23'],family['a24'],family['a25'],family['a26'],family['a27'],family['a28'],family['a29'],family['a30'],family['a31'],family['a32'],family['a33'],family['a34'],family['a35'],family['a36'],family['a37'],family['a38'],list_p,family['error'],family['state'],family['log']))
        if not(family['state']):
            if family['error']:
                error+=1
            start+=1
        else:
            end+=1
        n+=1
        print("已添加%s"% n)
    print("总共添加>>>>户数据<<<< %s项,%s项已完成，%s项待完成(其中%s项出错，待手工处理)"% (n,end,start,error))
    print("总共添加>>>>人数据<<<< %s项,%s项已完成，%s项待完成(其中%s项出错，待手工处理)"% (np,endp,startp,errorp))
    print('''
待录入数据准备完成，准备登陆系统...... GOGOGOGOGO

''')
    
#    print('''gogogo！！！按任意键开始启动Chrome浏览器......
#''')
#    input()

def error_json_to_xlsx(path,xlsxpath,list_js,list,error,start,end,n):
    '''将操作过的json文件筛查出错误的信息条，并写入到excel文件'''
    with open('002.json','r',encoding ="utf-8")as f:#加载json文件
        list_js =json.load(f)#将加载的json文件转换成字典列表
#input()

    for family in list_js:#将加载的json数据添加到personData列表中
        list.append(familyData.familyData(family['ID'],family['o1'],family['o2'],family['o3'],family['o4'],family['o5'],family['o6'],family['o7'],family['a1'],family['a2'],family['a3'],family['a4'],family['a5'],family['a6'],family['a7'],family['a8'],family['a9'],family['a10'],family['a11'],family['a12'],family['a13'],family['a14'],family['a15'],family['a16'],family['a17'],family['a18'],family['a19'],family['a20'],family['a21'],family['a22'],family['a23'],family['a24'],family['a25'],family['a26'],family['a27'],family['a28'],family['a29'],family['a30'],family['a31'],family['a32'],family['a33'],family['a34'],family['a35'],family['a36'],family['a37'],family['a38'],family['error'],family['state'],family['log']))
        if not(family['state']):
            if family['error']:
                error+=1
            start+=1
        else:
            end+=1
        n+=1
        
    print("总共%s项,%s项已完成，%s项待完成,其中%s项出错,按任意键导出到excel文件"% (n,end,start,error))
    input()

    for family in list_js:
        if not(family['state']):
            if family['error']:
                list.append(family)
                error+=1
            start+=1
        else:
            end+=1
        n+=1
    print("总共%s项,%s项已完成，%s项待完成,其中%s项出错,按任意键退出"% (n,end,start,error))
    input()




def personDatalist_to_json(familyDatalist,path):
    '''将family对象列表转化存储到json文件'''
    temp=[]
    person_list=[]

    with open(path,'r',encoding ="utf-8")as f:#加载json文件
        list_js =json.load(f)#将加载的json文件转换成字典列表
#input()

    for family in familyDatalist:
        for person in family.ps:
            person_list.append({'ID':person.ID,'o1':person.o1,'o2':person.o2,'o3':person.o3,'o4':person.o4,'o5':person.o5,'o6':person.o6,'b1':person.b1,'b2':person.b2,'b3':person.b3,'b4':person.b4,'b5':person.b5,'b6':person.b6,'b7':person.b7,'b8':person.b8,'b9':person.b9,'b10':person.b10,'b11':person.b11,'b12':person.b12,'b13':person.b13,'b14':person.b14,'b15':person.b15,'b16':person.b16,'b17':person.b17,'b18':person.b18,'b19':person.b19,'b20':person.b20,'b21':person.b21,'error':person.error,'state':person.state,'log':person.log})
        temp.append({'ID':family.ID,'o1':family.o1,'o2':family.o2,'o3':family.o3,'o4':family.o4,'o5':family.o5,'o6':family.o6,'o7':family.o7,'a1':family.a1,'a2':family.a2,'a3':family.a3,'a4':family.a4,'a5':family.a5,'a6':family.a6,'a7':family.a7,'a8':family.a8,'a9':family.a9,'a10':family.a10,'a11':family.a11,'a12':family.a12,'a13':family.a13,'a14':family.a14,'a15':family.a15,'a16':family.a16,'a17':family.a17,'a18':family.a18,'a19':family.a19,'a20':family.a20,'a21':family.a21,'a22':family.a22,'a23':family.a23,'a24':family.a24,'a25':family.a25,'a26':family.a26,'a27':family.a27,'a28':family.a28,'a29':family.a29,'a30':family.a30,'a31':family.a31,'a32':family.a32,'a33':family.a33,'a34':family.a34,'a35':family.a35,'a36':family.a36,'a37':family.a37,'a38':family.a38,'ps':person_list,'error':family.error,'state':family.state,'log':family.log})
    save_as_json(temp,path)
    print("已将处理后的情况保存到fp.json文件中，下次运行将直接读取进度")
    
def js_to_xlsx(js_path,xlsxpath):
    '''将缓存的json文件中错误的值'''
    with open(js_path,'r',encoding ="utf-8")as f:#加载json文件
        rec_data = json.load(f)#将加载的json文件转换成字典列表

    workbook = xlsxwriter.Workbook(xlsxpath)

    worksheet = workbook.add_worksheet()
    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
    worksheet2 =workbook.add_worksheet()

    # bold：加粗，num_format:数字格式

    bold_format = workbook.add_format({'bold': True})
    money_format = workbook.add_format({'num_format': '$#,##0'})
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15) 
    worksheet2.set_column(1, 1, 15)
    # 用符号标记位置，例如：A列1行
    n = 0
    np=0
    for key,value in rec_data[1].items():        
        worksheet.write(0,n, key, bold_format)
        n+=1
    #输入表头
    for key,value in rec_data[1]['ps'][1].items():        
        worksheet2.write(0,np, key, bold_format)
        np+=1
    
    row = 1
    countall = 0
    rowp = 1
    countallp = 0
    
    for item in rec_data:
        countall+=1
        for itemp in item['ps']:
            countallp+=1
            if itemp['error']:
                colp = 0               
                for key,value in itemp.items():
                    worksheet2.write_string(rowp, colp, str(value))
                    colp+=1
                rowp += 1
        # 使用write_string方法，指定数据格式写入数据
        if item['error']:
            col = 0               
            for key,value in item.items():
                worksheet.write_string(row, col, str(value))
                col+=1
            row += 1
    workbook.close()
    print("户数据总条数%d条，人数据总条数%d，提取户数据错误条数%d，人数据错误条数%d，文件保存在 %s 中。" % (countall ,countallp, (row - 1),(rowp - 1), xlsxpath))



def excel_json(path,path_two):  #第一步
    '''将excel文件数据转换为json并保存到本地'''
    wb = xlrd.open_workbook(path) 

    convert_list = []
    sh = wb.sheet_by_index(0)
    title = sh.row_values(0)
    for rownum in range(1, sh.nrows):
        rowvalue = sh.row_values(rownum)
        single = OrderedDict()
        for colnum in range(0, len(rowvalue)):
            #print(title[colnum], rowvalue[colnum])
            single[title[colnum]] = rowvalue[colnum]
        convert_list.append(single)
        print(rownum)
    j = str(json.dumps(convert_list,ensure_ascii=False))

    with codecs.open(path_two,"w",encoding='utf-8',errors='ignore') as f:
        f.write(j)
    print("excel转换json完成")

def js_to_xlsx_all(js_path,xlsxpath):
    '''将缓存的json文件分解成两张excel'''
    with open(js_path,'r',encoding ="utf-8")as f:#加载json文件
        rec_data = json.load(f)#将加载的json文件转换成字典列表

    workbook = xlsxwriter.Workbook(xlsxpath)

    worksheet = workbook.add_worksheet()
    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
    worksheet2 =workbook.add_worksheet()

    # bold：加粗，num_format:数字格式

    bold_format = workbook.add_format({'bold': True})
    money_format = workbook.add_format({'num_format': '$#,##0'})
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15) 
    worksheet2.set_column(1, 1, 15)
    # 用符号标记位置，例如：A列1行
    n = 0
    np=0
    for key,value in rec_data[1].items():        
        worksheet.write(0,n, key, bold_format)
        n+=1
    #输入表头
    for key,value in rec_data[1]['ps'][1].items():        
        worksheet2.write(0,np, key, bold_format)
        np+=1
    
    row = 1
    countall = 0
    rowp = 1
    countallp = 0
    
    for item in rec_data:
        countall+=1
        for itemp in item['ps']:
            countallp+=1
            if True:
                colp = 0               
                for key,value in itemp.items():
                    worksheet2.write_string(rowp, colp, str(value))
                    colp+=1
                rowp += 1
        # 使用write_string方法，指定数据格式写入数据
        if True:
            col = 0               
            for key,value in item.items():
                worksheet.write_string(row, col, str(value))
                col+=1
            row += 1
    workbook.close()
    print("户数据总条数%d条，人数据总条数%d，提取户数据错误条数%d，人数据错误条数%d，文件保存在 %s 中。" % (countall ,countallp, (row - 1),(rowp - 1), xlsxpath))










import sqlite3
import familyData
import personData
import json


def sql_insert(csor,id,key,value):
    pass

def json_sqlite(json,db,table):
    pass

def load_sql_family(csor,familynum):
    '''根据户编号查找符合的户数据'''
    sql="SELECT * FROM 'family' WHERE o3='%s'" % familynum
    
    csor.execute(sql)
    f = csor.fetchall()    
    familydata = familyData.familyData(*f[0])

    return familydata

def load_sql_person(csor,familynum):
    '''根据户编号查找符合的人数据'''
    sql2="SELECT * FROM 'person' WHERE o3='%s'" % familynum
    
    
    csor.execute(sql2)
    p = csor.fetchall()
    persons =[]
    for i in p:
        persons.append(personData.personData(*i))
    print("共加载%s条人数据"% len(persons))
    return persons

def loadfamilynums(csor):
    '''加载所有无错误和无完成状态的户编号'''
    sql = "select o3 from 'family' where error is null and state is null "
    csor.execute(sql)
    p =csor.fetchall()
    print("共加载%s条户数据"% len(p))
    return p  #返回的是[(a,),(b,)]



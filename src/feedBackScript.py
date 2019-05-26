import xlrd
import pandas as pd
from collections import Counter
from itertools import groupby
import re

'''
2019/5/14
根据业务割接计划自动生成割接反馈单
author：foxlora
'''
datadir = 'D:\\承载网业务割接\\2019.05.22\\'
filename = '2019年度业务系统割接计划（截止到5月24日）.xls'
outputFile = 'C:/Users/18351/Desktop/feedback.txt'
Asiafilter = ['徐龙飞']#['徐龙飞','李里','梁鸿洋']
Datefilter = 2       #设置1是周二，设置2是周四，设置其他为一周
BusSysNamefilter = ['NE40E','PCRP1','NE80E','X16','X8A']


class FeedBack(object):
    def __init__(self):
        self.date = None#读取割接日期
        self.VPNNum = None#读取VPN网管工单号
        self.Provinces = None #省份
        self.Cities = None #地市
        self.Manufactor = None  #厂家
        self.BusinessName = None #业务名称
        self.BusinessType = None #业务类型
        self.BussSysName = None #业务系统设备名称
        self.AsiainfoContacts = None#亚信联系人

    #读取割接反馈单信息
    def readexcel(self):
        excelname = datadir + filename
        self.data = xlrd.open_workbook(excelname)#读取表格文件

        self.table = self.data.sheets()[0]#读取第二个sheet内容

        #获取sheet行数，列数
        print('行数:',self.table.nrows)
        print('列数:',self.table.ncols)
        #获取整列值
        self.date = self.table.col_values(1)#读取割接日期
        self.VPNNum = self.table.col_values(4)#读取VPN网管工单号
        self.Provinces = self.table.col_values(5) #省份
        self.Cities = self.table.col_values(6) #地市
        self.Manufactor = self.table.col_values(7)  #厂家
        self.BusinessName = self.table.col_values(8) #业务名称
        self.BusinessType = self.table.col_values(9) #业务类型
        self.BussSysName = self.table.col_values(11) #业务系统设备名称
        self.AsiainfoContacts = self.table.col_values(16)#亚信联系人


        #获得频率最高的两个日期
        counter = Counter(self.date)
        date_two = counter.most_common()

        firstDateIndex = self.listValueToIndex(self.date,date_two[0][0])
        secondDateIndex = self.listValueToIndex(self.date,date_two[1][0])



        #firstDateRows = [i for i in firstDateIndex if '徐龙飞' in AsiainfoContacts[i]]
        firstDateRows = [i for i in firstDateIndex for f in Asiafilter if f in self.AsiainfoContacts[i] ]
        secondDateRows = [i for i in secondDateIndex for f in Asiafilter if f in self.AsiainfoContacts[i] ]

        print(firstDateRows)
        print(secondDateRows)
        if Datefilter == 1:
            self.DateRows = firstDateRows
        elif Datefilter == 2:
            self.DateRows = secondDateRows
        else:
            self.DateRows = firstDateRows + secondDateRows
            print(self.DateRows)


    #建立工单号字典{'GD-20190508103756327'：[51,52]}
    def makeDict(self):
        self.vpndict = {}
        for i in self.DateRows:
            if self.VPNNum[i] in self.vpndict:
                self.vpndict[self.VPNNum[i]].append(i)
            else:
                self.vpndict[self.VPNNum[i]] = [i]



    def businessProcess(self):
        #处理每一个业务工单
        for key in self.vpndict.keys():
            pairList = self.vpndict[key]
            resultnumlist = self.GroupbyVPNtype(pairList)
            ResultRowList = self.RowsCompare(resultnumlist)
            self.vpnMakeScript(ResultRowList)


    def vpnMakeScript(self,ResultRowList):
        for vpnRows in ResultRowList:
            rowsnum = len(vpnRows)#计算多少行
            k,v = divmod(rowsnum,2)
            if k:
                for i in range(k):
                    self.make_script(vpnRows[2*i],vpnRows[2*i + 1])
            if v:
                self.make_script_singlerow(vpnRows[-1])





    #根据不同的业务类型、不同的VPN返回列表
    def GroupbyVPNtype(self,pairlist):
        vpntype = {}
        resultlist = []
        #[[57],[58,59,60,61]]
        for i in pairlist:
            if (self.BusinessName[i],self.BusinessType[i]) in vpntype:
                vpntype[(self.BusinessName[i],self.BusinessType[i])].append(i)
            else:
                vpntype[(self.BusinessName[i],self.BusinessType[i])] = [i]
        for i in vpntype.values():
            resultlist.append(i)
        return resultlist





    #根据value获取list的值索引，返回索引的list
    def listValueToIndex(self,linklist,val):
        indexList = []
        for index,value in enumerate(linklist):
            if value == val:
                indexList.append(index)
        return indexList





    def readexcel_pandas(self):
        excelname = datadir + filename
        data = pd.read_excel(excelname)#读取表格文件
        print(data)


    #行比较
    def RowsCompare(self,resultnumlist):
        ResultRowList = []
        for rownumList in resultnumlist:#[57,58]
            rowList = [self.table.row_values(i) for i in rownumList]
            rowList.sort(key=lambda x:x[11])
            ResultRowList.append(rowList)
        return ResultRowList


    def make_script(self,row1,row2):

        province = row1[5]
        city = row1[6]
        manufactor = row1[7]
        businessName = row1[8] #业务名称
        businessType = row1[9] #业务类型
        bussSysName1 = row1[11] #业务系统设备名称
        bussSysName2 = row2[11]
        try:
            cenum1 = re.findall(r'CE(\d+)', bussSysName1)[0]
            cenum2 = re.findall(r'CE(\d+)', bussSysName2)[0]

        except Exception as e:
            for filterstr in BusSysNamefilter:
                bussSysName1 = bussSysName1.replace(filterstr, '')
                bussSysName2 = bussSysName2.replace(filterstr, '')
            cenum1 = re.sub('\D', '', bussSysName1)
            cenum2 = re.sub('\D', '', bussSysName2)

        finally:
            if len(cenum1) < 2:
                cenum1 = '0' + cenum1
            if len(cenum2) < 2:
                cenum2 = '0' + cenum2

            script = "{0} {1} {2} {3} {4} CE{5}、CE{6}\n"
            script = script.format(province, city, manufactor, businessName, businessType, cenum1, cenum2)
            with open(outputFile, 'a+') as f:
                f.write(script)







    def make_script_singlerow(self,row1):
        province = row1[5]
        city = row1[6]
        manufactor = row1[7]
        businessName = row1[8] #业务名称
        businessType = row1[9] #业务类型
        bussSysName1 = row1[11] #业务系统设备名称

        try:
            cenum1 = re.findall(r'CE(\d+)', bussSysName1)[0]
        except:

            for filterstr in BusSysNamefilter:
                bussSysName1 = bussSysName1.replace(filterstr,'')

            cenum1 = re.sub('\D','',bussSysName1)
        finally:
            if len(cenum1)< 2:
                cenum1 = '0' + cenum1

            script = "{0} {1} {2} {3} {4} CE{5}\n"
            script = script.format( province, city, manufactor,businessName,businessType,cenum1)
            with open(outputFile, 'a+') as f:
                f.write(script)

if __name__ == '__main__':
    FB = FeedBack()
    FB.readexcel()
    FB.makeDict()
    FB.businessProcess()



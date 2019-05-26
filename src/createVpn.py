import xlrd
from collections import Counter
import re
'''
2019/5/22
根据新增VPN调单自动生成脚本
'''
datadir = 'D:\\承载网业务割接\\2019.05.24\\'
filename = '20190542吉林(PS-IMS)调度单.xls'
outputFile = 'C:/Users/18351/Desktop/createVpn.txt'



class CreateVPN(object):
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

    #读取调单信息
    def readexcel(self):
        excelname = datadir + filename
        self.data = xlrd.open_workbook(excelname)#读取表格文件
        self.table = self.data.sheets()[1]#读取第二个sheet内容
        self.vpnName = self.table.row_values(3)[1]
        self.vpn_RT = self.table.row_values(3)[14]
        self.vpn_RD = self.table.row_values(3)[17]
        # print(self.data.sheet_names())
        #获取sheet行数，列数
        print('行数:',self.table.nrows)
        print('列数:',self.table.ncols)

        #获取整列值
        self.CEDeviceName = self.table.col_values(4)#CE设备名称
        self.CEport = self.table.col_values(9)#CE与PE设备互联端口
        self.Routes = self.table.col_values(14) #承载网PE对外发布的VPN路由
        self.BandWidth = self.table.col_values(15) #需求带宽(M)
        self.PEDeviceName = self.table.col_values(16)  #PE设备名称
        self.PEport = self.table.col_values(19) #PE与CE设备互联端口
        self.PEipAddress = self.table.col_values(21) #PE与CE设备互联端口地址
        self.BusinessType = self.table.col_values(25) #业务接入类型
        print(self.PEport)



    def vpninstance_Generation(self,vpnName,IPv4OR6,RD,RT):
        vpn_script = 'ip vpn-instance {vpnName}\n' \
                     ' {ipvx}-family\n' \
                     '  route-distinguisher {RD}\n' \
                     '  apply-lable per-instance\n' \
                     '  routing-table limit 100000 80\n' \
                     '  vpn-target {RT} export-extcommunity\n' \
                     '  vpn-target {RT} import-extcommunity\n'
        vpn_script = vpn_script.format(vpnName=vpnName,ipvx=IPv4OR6,RD=RD,RT=RT)
        with open(outputFile, 'a+') as f:
            f.write(vpn_script)


    def physical_port_Gen(self,PEport,SubPort,BusName,CEname,CEport,BandWidth,VpnName,PEipAdd):
        port_script = '2'



    def make_script_singlerow(self,row1):
        with open(outputFile, 'a+') as f:
            f.write()

if __name__ == '__main__':
    CV = CreateVPN()
    CV.readexcel()
    #CV.vpninstanceGeneration('ChinaMobile_IMS_SG','ipv4','24059:66010','24059:66010')




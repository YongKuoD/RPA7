#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023/4/18 20:53
# @Author  : YongKuo
# @Email   : YongKuoD@gmail.com
# @File    : RPA7.py
# @Project : rpa1
# @Software: PyCharm


import requests
import json
import pandas as pd
import os
import datetime

# excel 表格规范字段，和表格表头一一对应，顺序不可更改
excelKey = ['invoiceCode', 'invoiceNo', 'invoiceHeader', 'taxId', 'registerAddressPhone', 'bankAccountNumber',
                    'contactMobile', 'email',
                    'sellerName',  'invoiceType',
                    'invoiceType',  'descr', 'deviceType', 'invoiceListMark', 'serialNo',
                    'goodsCode', 'goodsName', 'goodsTaxRate', 'goodsSpecification',
                    'goodsUnit', 'goodsQuantity', 'includTaxgoodsPrice', 'goodsPrice', 'priceTaxMark',
                    'includTaxgoodsTotalPrice',
                    'goodsPrice', 'goodsTotalTax', 'invoiceTotalPricelncludeTax', 'invoiceTotalPrice',
                    'invoiceTotalTax', 'invoiceTotalPriceTax', 'invoiceUploadMark', 'invoiceDate',
                    'invoiceStatus', 'invoiceInvalidDate','machineCode', 'orderNo',  'sourceMark',
                    'invoiceCheckMark', "playStatus", 'invoiceStatus', "className", "studentName"]

notkeys = ['deductibleAmount','invoiceInvalidDate','invoiceInvalidDate']
# 获取 excel 表格表头规范
def get_clomns():
    levels = [['发票代码','发票号码','购货单位名称','购货纳税人识别号','购货单位地址电话','购货单位银行账号','购方客户电话','购方客户邮箱',
               '销货单位名称','发票类型','开票类型',
               '备注','设备类型','清单标志','发票请求流水号','发票明细','合计金额（含税）','合计金额（不含税）',
               '合计税额','价税合计','开票终端标识','开票日期','发票状态','作废日期','机器编号','业务发票请求流水号','来源标识',
               '验签状态','支付状态','一体机开票状态','班级名称（备注）','教务系统学员姓名'],
                ['','商品编码','商品名称' ,'税率','规格型号','单位','数量','单价（含稅）','单价（不含稅）',
                '含税标志','金额（含税）','金额（不含税）','税额']]

    codes=[[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,15,15,15,15,15,15,15,15,15,15,15,16,17,18,19,
            20,21,22,23,24,25,26,27,28,29,30,31],
            [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,2,3,4,5,6,7,8,9,10,11,12,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]]

    mulClomns = pd.MultiIndex(levels=levels,
                   codes=codes)

    return mulClomns



class rpa1(object):

    def __init__(self):
        self.get_config()


    def get_config(self):
        # print(os.getcwd())
        configFile = os.path.join(os.getcwd(), "配置文件.xlsx")
        dataframe = pd.read_excel(configFile,header=None)
        config = list(dataframe.loc[:, 1])
        self.targetDir = config[0]
        if not os.path.exists(self.targetDir):
            os.makedirs(self.targetDir)
        self.stime = config[1]
        self.etime = config[2]

    def get_request(self,url):
        response = requests.get(url=url)
        content = json.loads(response.text)
        return content

    # 获取第一个接口的数据
    def get_data_leve1_1(self):
        '''
                    购货单位名称	 invoiceHeader
                    购货纳税人识别号	taxId
                    购货单位地址	registerAddress
                    购货单位电话	registerMobile
                    购货单位银行账号	bankAccountNumber
                    购方客户电话 	contactMobile
                    购方客户邮箱	email
                    发票类型		invoiceType
                    支付状态		playStatus
                    班级名称（备注）	className
                    备注			descr
                    教务系统学员姓名	studentName
                    unid			unid
        :return:
        '''
        urlLevel1 = "http://60.205.245.225:31315/cms/web/invoiceInfo?startTime=%s 00:00:00&endTime=%s 23:59:59"%(self.stime,self.etime)
        data_L1 = self.get_request(urlLevel1)["data"]
        if not data_L1:
            return

        keys = ['invoiceHeader','taxId','registerAddress','registerMobile','bankAccountNumber','contactMobile',
                'email','invoiceType','playStatus','className','descr','studentName','unid', 'invoiceStatus',
                'machineCode']
        dataList = []
        for data in data_L1:
            dataDict = {}
            for k in keys:
                dataDict[k] = data[k]
            dataDict['className'] = dataDict['className'].replace(dataDict['studentName'],'')
            unid = data['unid']
            dict_L2 = self.get_data_leve1_2(unid)
            dict_L3 = self.get_data_leve1_3(unid)
            dataDict.update(dict_L2)
            dataDict.update(dict_L3)
            dataList.append(dataDict)
        dataFrame = pd.DataFrame(dataList)
        # 加法 被加数: [  # x.goodsTotalPrice 加数: (#xgoodsTotalTax] 存储结果到变量: (#fx.goodsTotalPricelncludeTax
        dataFrame['goodsTotalPricelncludeTax'] = dataFrame['goodsTotalTax'] + dataFrame['goodsTotalPrice']
        # 加法: 合计金额(含税) 被加数: [  # fx.invoiceTotalTax 加数: [#fxinvoiceTotalPrice] 存储结果到变量: [xinvoiceTotalPricelncludeTax
        dataFrame['invoiceTotalPricelncludeTax'] = dataFrame['invoiceTotalTax'] + dataFrame['invoiceTotalPrice']
        # 地址电话拼接
        dataFrame['registerAddressPhone'] = dataFrame['registerAddress'] + dataFrame['registerMobile']

        for kn in notkeys:
            dataFrame[kn] = ''

        dataFrame = dataFrame[excelKey]
        dataFrame.columns = get_clomns()
        return dataFrame



    def get_data_leve1_2(self,unid):
        '''
            合计税额		invoiceTotalTax
            价税合计		invoiceTotalPriceTax
            合计金额（不含税）invoiceTotalPrice
            开票日期		invoiceDate
            发票代码		invoiceCode
            发票号码		invoiceNo
            发票行行号		goodsLineNo
            发票请求流水号	serialNo
            商品编码		goodsCode
            商品名称		goodsName
            税率			goodsTaxRate
            含税标志		priceTaxMark
            金额（不含税）	goodsTotalPrice
            数量			goodsQuantity
            单位			goodsUnit
            单价（不含税）	goodsPrice
            税额			goodsTotalTax
            一体机开票状态	invoiceStatus
        :param unid:
        :return:
        '''
        baseKeys = ['invoiceTotalTax', 'invoiceTotalPriceTax', 'invoiceTotalPrice', 'invoiceDate','invoiceCode','invoiceNo',
                'serialNo']
        detailsKeys =['goodsLineNo','goodsCode','goodsName','goodsTaxRate','priceTaxMark','goodsTotalPrice',
                'goodsQuantity', 'goodsUnit', 'goodsPrice', 'goodsTotalTax']

        urlLeverl2 = "http://60.205.245.225:31315/cms/web/invoiceIssue?unid=%s"%(unid)

        # requestsData = self.get_request(urlLeverl2)
        try :
            data_L2 = self.get_request(urlLeverl2)['data']['response']['success'][0]
        except:
            dataDict = {}
            for k in (baseKeys+ (detailsKeys)):
                dataDict[k] = None
        else:
            detailsListData = data_L2['invoiceDetailsList'][0]

            dataDict = {}
            for kb in baseKeys:
                dataDict[kb] = data_L2[kb]
            for kd in detailsKeys:
                dataDict[kd] = detailsListData[kd]

        return dataDict


    def get_data_leve1_3(self,unid):
        '''
            设备类型		deviceType
            清单标志		invoiceListMark
            征收方式		taxationMethod
            签验状态		invoiceCheckMark
            机构税号		sellerTaxNo
            销货单位名称	sellerName
            销货单位纳税识别号	buyerTaxNo
            销货单位地址电话		buyerAddressPhone
            销货单位银行账号		buyerBankAccount
            收款人			payee
            审核人			checker
            开票人			drawer
        :param unid:
        :return:
        '''

        baseKeys = ['deviceType', 'invoiceListMark', 'taxationMethod','invoiceCheckMark', 'sellerTaxNo','sellerName',
                'buyerTaxNo','buyerAddressPhone', 'payee','checker','drawer','invoiceUploadMark','orderNo',
                'sourceMark','buyerBankAccount']
        detailsKeys = ['includTaxgoodsTotalPrice','includTaxgoodsPrice','goodsSpecification','preferentialMarkFlag',
                       'invoiceLineNature']
        urlLeverl3 = "http://60.205.245.225:31315/cms/web/invoiceQuery?unid=%s"%(unid)
        try:
            data_L3 = self.get_request(urlLeverl3)['data'][0]['response'][0]
        except:
            dataDict = {}
            for k in (baseKeys+detailsKeys):
                dataDict[k] = None
        else:
            dataDict = {}
            # invoiceDetailsList
            detailsData = data_L3['invoiceDetailsList'][0]
            for kd in detailsKeys:
                dataDict[kd] = detailsData[kd]
            for kb in baseKeys:
                dataDict[kb] = data_L3[kb]

        return dataDict

    # 判断 dataframe 是否为空
    def is_empty(self,dataframe):
        if isinstance(dataframe, pd.DataFrame):
            return dataframe.empty
        else:
            return True

    def create(self):
        dataFrame = self.get_data_leve1_1()

        if self.is_empty(dataFrame):
           return
        print(dataFrame)
        dateTime = str(datetime.date.today())
         #开票所有信息表
        allDataFile = os.path.join(self.targetDir,"开票所有信息表.xlsx")
        dataFrame.index = list(range(1, dataFrame.shape[0] + 1))
        dataFrame.to_excel(allDataFile, index_label='序号')
#       班级开票信息
        className = set(dataFrame['班级名称（备注）'])
        classDataDir = os.path.join(self.targetDir,dateTime)
        if not os.path.exists(classDataDir):
            os.makedirs(classDataDir)

        classDataFile = os.path.join(classDataDir,'班级开票信息表.xlsx')
        writer = pd.ExcelWriter(classDataFile)


        for name in className:

            classDataframe = dataFrame[dataFrame['班级名称（备注）'] == name]
            if len(name) > 31:
                name = name[:31]
            classDataframe.index = list(range(1, classDataframe.shape[0] + 1))
            classDataframe.to_excel(writer, sheet_name=name, index=True,
                                                 index_label='序号')
        writer._save()




if __name__ == "__main__":

    rpa = rpa1()
    rpa.create()



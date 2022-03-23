import openpyxl
import pandas as pd
import os
import matplotlib as plt


##包含第一层统计结果
class SteelStat(object):
    def __init__(self,dir):
        self.dir=dir
        
               
    def readExcel(self):
        sheetname="钢筋工程量"
        return pd.read_excel(self.dir,sheet_name=sheetname,header=None)
    
    def cleanExcel(self):
        data=self.readExcel()
        #表格的第一行nan内容替换为前置
        data.loc[0,:]=data.loc[0,:].fillna(method='ffill')
        #表格的第一第二列nan内容替换为前置
        data.loc[:,[0,1]]=data.loc[:,[0,1]].fillna(method='ffill')
        #将数据部分nan替换为0
        data=data.fillna(value=0)
        #不需要的行删掉
        data=data[~data.loc[:,1].str.contains('合计')]  
        return data
    #将梁的数据提取出来
    def beamSt(self):
        data=self.cleanExcel()
        #复制一份数据
        df=data[data.loc[:,0].str.contains('梁')].copy()
        #只要float部分
        df.loc[:,2:]=df.loc[:,2:].astype(float)
        #将求和结果作为新行添加到最后一行
        df.loc['N']=df.loc[:,2:].sum(axis=0)
        #每层含钢量计算后作为新一列插入dataframe，保留两位小数
        df.iloc[:,-1]=(df.iloc[:,-2]/df.iloc[:,2]).astype(float).round(2)
        df=df.fillna("合计")
        return df
        #柱
    def coluSt(self):
        data=self.cleanExcel()
        df=data[data.loc[:,0].str.contains('柱')].copy()
        df.loc[:,2:]=df.loc[:,2:].astype(float)
        df.loc['N']=df.loc[:,2:].sum(axis=0)
        df.iloc[:,-1]=(df.iloc[:,-2]/df.iloc[:,2]).astype(float).round(2)
        df=df.fillna("合计")
        return df
        #板
    def slabSt(self):
        data=self.cleanExcel()
        df=data[data.loc[:,0].str.contains('板')].copy()
        df.loc[:,2:]=df.loc[:,2:].astype(float)
        df.loc['N']=df.loc[:,2:].sum(axis=0)
        df.iloc[:,-1]=(df.iloc[:,-2]/df.iloc[:,2]).astype(float).round(2)
        df=df.fillna("合计")
        return df
        #墙
    def wallSt(self):
        data=self.cleanExcel()
        df=data[data.loc[:,0].str.contains('墙')].copy()
        df.loc[:,2:]=df.loc[:,2:].astype(float)
        df.loc['N']=df.loc[:,2:].sum(axis=0)
        df.iloc[:,-1]=(df.iloc[:,-2]/df.iloc[:,2]).astype(float).round(2)
        df=df.fillna("合计")
        return df
    #数据拼接
    def resultSt(self):
        data=self.cleanExcel()
        df1=self.beamSt()
        df2=self.coluSt()
        df3=self.slabSt()
        df4=self.wallSt()
        df5=data[data.loc[:,0].str.contains('类别')]
        frames=[df5,df1,df2,df3,df4]
        #将数据合并
        data2=pd.concat(frames,ignore_index=True)
        data2.loc['n']=data2[data2.loc[:,0].str.contains('合计')].loc[:,2:].astype(float).sum(axis=0)
        data2=data2.fillna("合计")
        return data2
    #绘制条形图查看梁柱板墙的用钢量
    def drawBar(self):
        pass
    #保存为结果表格
    def saveExcel(self):
        data=self.resultSt()
        data.to_excel('Result3.xlsx','sheet1',index=0,header=0)
        
        
#不包含第一层的统计        
class SteelStat1(SteelStat):
    def __init__(self,dir):
        super().__init__(dir)
        
    #将子类的cleanExcel方法重写，其余的都继承于SteelStat
    def cleanExcel(self):
        data=self.readExcel()
        #表格的第一行nan内容替换为前置
        data.loc[0,:]=data.loc[0,:].fillna(method='ffill')
        #表格的第一第二列nan内容替换为前置
        data.loc[:,[0,1]]=data.loc[:,[0,1]].fillna(method='ffill')
        #将数据部分nan替换为0
        data=data.fillna(value=0)
        #不需要的行删掉
        data=data[~data.loc[:,1].str.contains('合计|第1层')]  #多值采用|运算
        return data
    
        
data=SteelStat1(dir=r"D:\GHPython\pandas\test.xlsx")
data.saveExcel()
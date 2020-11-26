# coding=<utf-8>
from datetime import date, datetime
import pandas as pd
import os
# from pandas.io import spss
import xlsxwriter
# from dateutil.parser import parse

#----------待完成----------
#计算入职月
#计算考核期
#入职首月业绩划分
#当前考核期业绩汇总

print('正在处理...')

#----------获取文件名----------
filename = os.walk('./')
xlsxname = []
for a,b,c in filename:   
    xlsxname = c
# print(type(xlsxname))
listlen = len(xlsxname)

tempnum = 0 
xlsxnum = 0
while tempnum < listlen:
    if ('.xlsx' in xlsxname[tempnum]) and ('公司月度业绩明细报表（个寿）' in xlsxname[tempnum]):
        xlsxnum = tempnum
    tempnum = tempnum + 1

#----------读取文件----------
DataImport = pd.read_excel(xlsxname[xlsxnum])
# DataImport = pd.read_excel(r'.\data.xlsx')  # 导入初始数据
DataInit = DataImport
#DataInit = DataImport.drop(DataImport.columns[[2, 3, 4, 5, 8, 9, 10, 11, 12, 14, 17,18, 19, 20, 21, 22, 24, 25, 28, 29, 30, 31, 32, 33, 34, 37, 39]], axis=1)  # 删除不需要的行

#----------数据预处理----------
DataInit = DataInit.astype({' 缴费方式 ': 'string', ' 服务人员 ': 'string', ' 服务人员工号 ': 'string', ' 保单状态 ': 'string',
                            ' 是否医疗险 ': 'string', ' 医疗部产品分类 ': 'string', ' 分单人类型 ': 'string'})  # object转换string
DataInit[' 交单日期 '] = pd.to_datetime(DataInit[' 交单日期 '])  # object转换datetime
DataInit[' 回执操作日期 '] = pd.to_datetime(DataInit[' 回执操作日期 '])  ## object转换datetime

# Data1 业绩追踪
# Data2 MDRT追踪

# Data1 选择正式保单,去掉辅助交单
Data1 = DataInit[DataInit[' 保单状态 '].str.contains('正式保单')]
Data1 = Data1.fillna({' 分单人类型 ': '非联合交单'})
Data1 = Data1.astype({' 分单人类型 ': 'string'})
Data1 = Data1[~Data1[' 分单人类型 '].str.contains('辅助交单人')]

#选择正式保单,等待回执,投保单,去掉辅助交单
Data2 = DataInit[DataInit[' 保单状态 '].str.contains('正式保单|等待回执|投保单')]
Data2 = Data2.fillna({' 分单人类型 ': '非联合交单'})
Data2 = Data2.astype({' 分单人类型 ': 'string'})
Data2 = Data2[~Data2[' 分单人类型 '].str.contains('辅助交单人')]

#Data1 = Data1.drop(Data1.columns[[0, 1, 2, 4, 6, 9, 10, 11, 12]], axis=1)
#Data1 = Data1.insert()
DataMonPerf = pd.DataFrame(columns=['2020-01月业绩', '2020-02月业绩', '2020-03月业绩', '2020-04月业绩',
                                    '2020-05月业绩', '2020-06月业绩', '2020-07月业绩', '2020-08月业绩', '2020-09月业绩', '2020-10月业绩', '2020-11月业绩', '2020-12月业绩'])
Data1 = pd.concat([Data1, DataMonPerf])

DataMDRT = pd.DataFrame(columns=['折算保费','2020FYC','是否为2020业绩'])
Data2 = pd.concat([Data2,DataMDRT])

#----------计算每月业绩----------
Data1Shape = list(Data1.shape)
Data1RowNum = Data1Shape[0]  # Data1行数

Month1912EndReceiptDate = datetime(2020, 2, 14)

Month2001StartOrderDate = datetime(2020, 1, 23)
Month2001EndOrderDate = datetime(2020, 2, 25)
Month2001EndReceiptDate = datetime(2020, 3, 10)

Month2002StartOrderDate = datetime(2020, 2, 26)
Month2002EndOrderDate = datetime(2020, 3, 25)
Month2002EndReceiptDate = datetime(2020, 4, 10)

Month2003StartOrderDate = datetime(2020, 3, 26)
Month2003EndOrderDate = datetime(2020, 4, 27)
Month2003EndReceiptDate = datetime(2020, 5, 12)

Month2004StartOrderDate = datetime(2020, 4, 28)
Month2004EndOrderDate = datetime(2020, 5, 25)
Month2004EndReceiptDate = datetime(2020, 6, 10)

Month2005StartOrderDate = datetime(2020, 5, 26)
Month2005EndOrderDate = datetime(2020, 6, 24)
Month2005EndReceiptDate = datetime(2020, 7, 10)

Month2006StartOrderDate = datetime(2020, 6, 25)
Month2006EndOrderDate = datetime(2020, 7, 27)
Month2006EndReceiptDate = datetime(2020, 8, 10)

Month2007StartOrderDate = datetime(2020, 7, 28)
Month2007EndOrderDate = datetime(2020, 8, 25)
Month2007EndReceiptDate = datetime(2020, 9, 10)

Month2008StartOrderDate = datetime(2020, 8, 26)
Month2008EndOrderDate = datetime(2020, 9, 25)
Month2008EndReceiptDate = datetime(2020, 10, 12)

Month2009StartOrderDate = datetime(2020, 9, 26)
Month2009EndOrderDate = datetime(2020, 10, 26)
Month2009EndReceiptDate = datetime(2020, 11, 12)

Month2010StartOrderDate = datetime(2020, 10, 27)
Month2010EndOrderDate = datetime(2020, 11, 25)
Month2010EndReceiptDate = datetime(2020, 12, 10)

Month2011StartOrderDate = datetime(2020, 11, 26)
Month2011EndOrderDate = datetime(2020, 12, 25)
Month2011EndReceiptDate = datetime(2021, 1, 11)

Month2012StartOrderDate = datetime(2020, 12, 26)
Month2012EndOrderDate = datetime(2021, 1, 25)
Month2012EndReceiptDate = datetime(2021, 2, 9)

#1月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2001StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2001EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2001EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2001StartOrderDate) and (Month1912EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2001EndReceiptDate)):
        Data1.iat[PerfNum,40] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,40] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#2月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2002StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2002EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2002EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2002StartOrderDate) and (Month2001EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2002EndReceiptDate)):
        Data1.iat[PerfNum,41] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,41] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#3月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2003StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2003EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2003EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2003StartOrderDate) and (Month2002EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2003EndReceiptDate)):
        Data1.iat[PerfNum,42] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,42] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#4月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2004StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2004EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2004EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2004StartOrderDate) and (Month2003EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2004EndReceiptDate)):
        Data1.iat[PerfNum,43] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,43] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#5月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2005StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2005EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2005EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2005StartOrderDate) and (Month2004EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2005EndReceiptDate)):
       Data1.iat[PerfNum,44] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,44] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#6月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2006StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2006EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2006EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2006StartOrderDate) and (Month2005EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2006EndReceiptDate)):
        Data1.iat[PerfNum,45] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,45] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#7月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2007StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2007EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2007EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2007StartOrderDate) and (Month2006EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2007EndReceiptDate)):
        Data1.iat[PerfNum,46] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,46] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#8月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2008StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2008EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2008EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2008StartOrderDate) and (Month2007EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2008EndReceiptDate)):
        Data1.iat[PerfNum,47] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,47] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#9月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2009StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2009EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2009EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2009StartOrderDate) and (Month2008EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2009EndReceiptDate)):
        Data1.iat[PerfNum,48] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,48] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#10月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2010StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2010EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2010EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2010StartOrderDate) and (Month2009EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2010EndReceiptDate)):
        Data1.iat[PerfNum,49] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,49] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#11月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2011StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2011EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2011EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2011StartOrderDate) and (Month2010EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2011EndReceiptDate)):
        Data1.iat[PerfNum,50] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,50] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)

#12月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2012StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2012EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2012EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2012StartOrderDate) and (Month2011EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2012EndReceiptDate)):
        Data1.iat[PerfNum,51] = Data1.iloc[PerfNum,7]
    else:
        Data1.iat[PerfNum,51] = 0
    PerfNum = PerfNum + 1
# print(PerfNum)


Data1Pivot = pd.pivot_table(Data1,values=['2020-01月业绩', '2020-02月业绩', '2020-03月业绩', '2020-04月业绩',
                                    '2020-05月业绩', '2020-06月业绩', '2020-07月业绩', '2020-08月业绩', '2020-09月业绩', '2020-10月业绩', '2020-11月业绩', '2020-12月业绩'],index=' 服务人员 ',aggfunc='sum')


#计算MDRT折算保费,FYC
Data2Shape = list(Data2.shape)
Data2RowNum = Data2Shape[0]  # Data1行数
PerfNum = 0
while PerfNum < Data2RowNum:
    if (((Month2001StartOrderDate <= Data2.iloc[PerfNum,23]) and (Data2.iloc[PerfNum, 23] <= Month2012EndOrderDate) and (Data2.iloc[PerfNum, 26] <= Month2012EndReceiptDate)) or ((Data2.iloc[PerfNum, 23] < Month2001StartOrderDate) and (Month1912EndReceiptDate < Data2.iloc[PerfNum, 26]) and (Data2.iloc[PerfNum,26] <= Month2012EndReceiptDate))) :
        Data2.iat[PerfNum,42] = 1
    else:
        Data2.iat[PerfNum,42] = 0 
    if ((Data2.iloc[PerfNum,27] == '等待回执') or (Data2.iloc[PerfNum,27] == '投保单')):
        Data2.iat[PerfNum,42] = 1
    PerfNum = PerfNum + 1

PerfNum = 0
while PerfNum < Data2RowNum:
    if Data2.iloc[PerfNum,42] == 1 :
        Data2.iat[PerfNum,40] = Data2.iloc[PerfNum,1]
        Data2.iat[PerfNum,41] = Data2.iloc[PerfNum,6]
    else:
        Data2.iat[PerfNum,40] = 0
        Data2.iat[PerfNum,41] = 0
    if (Data2.iloc[PerfNum,42] == 1 and (Data2.iloc[PerfNum,36] == '普通个寿')) and (Data2.iloc[PerfNum,14] == '1年'):
        Data2.iat[PerfNum,40] = 0.06 * Data2.iloc[PerfNum,1]
    PerfNum = PerfNum + 1

Data2Pivot = pd.pivot_table(Data2,values=['折算保费', '2020FYC'],index=' 服务人员 ',aggfunc='sum')

Data2Pivot = Data2Pivot.sort_values(by = ["折算保费"],ascending = False)
# Data2Pivot['MDRT-FYC差值'] = ''
# Data2Pivot['MDRT-保费差值'] = ''
# Data2Pivot['COT-FYC差值'] = ''
# Data2Pivot['COT-保费差值'] = ''
# Data2Pivot['TOT-FYC差值'] = ''
# Data2Pivot['TOT-保费差值'] = ''
# DataMDRT2 = pd.DataFrame(columns=['MDRT-FYC差值','MDRT-保费差值','COT-FYC差值','COT-保费差值','TOT-FYC差值','TOT-保费差值'])
# Data2Pivot = pd.concat([Data2Pivot,DataMDRT2])


# Data2PivotShape = list(Data2Pivot.shape)
# Data2PivotRowNum = Data2PivotShape[0]  # Data1行数
# PerfNum = 0

# while PerfNum < Data2PivotRowNum:
#     if(Data2Pivot.iloc[PerfNum, 1] < 171300):
#         Data2Pivot.iat[PerfNum, 3] = 171300 - Data2Pivot.iloc[PerfNum, 1]
#     else:
#         Data2Pivot.iat[PerfNum, 3] = '预达成'

#     if(Data2Pivot.iloc[PerfNum, 2] < 513900):
#         Data2Pivot.iat[PerfNum, 4] = 513900 - Data2Pivot.iloc[PerfNum, 2]
#     else:
#         Data2Pivot.iat[PerfNum, 4] = '预达成'

#     if(Data2Pivot.iloc[PerfNum, 1] < 513900):
#         Data2Pivot.iat[PerfNum, 5] = 513900 - Data2Pivot.iloc[PerfNum, 1]
#     else:
#         Data2Pivot.iat[PerfNum, 5] = '预达成'

#     if(Data2Pivot.iloc[PerfNum, 2] < 1541700):
#         Data2Pivot.iat[PerfNum, 6] = 1541700 - Data2Pivot.iloc[PerfNum, 2]
#     else:
#         Data2Pivot.iat[PerfNum, 6] = '预达成'

#     if(Data2Pivot.iloc[PerfNum, 1] < 1027800):
#         Data2Pivot.iat[PerfNum, 7] = 1027800 - Data2Pivot.iloc[PerfNum, 1]
#     else:
#         Data2Pivot.iat[PerfNum, 7] = '预达成'

#     if(Data2Pivot.iloc[PerfNum, 2] < 3083400):
#         Data2Pivot.iat[PerfNum, 8] = 3083400 - Data2Pivot.iloc[PerfNum, 2]
#     else:
#         Data2Pivot.iat[PerfNum, 8] = '预达成'
#     PerfNum = PerfNum + 1
# Data2Pivot = pd.concat([Data2Pivot,DataMDRT2])
# Data2Pivot = pd.concat([Data2Pivot,DataMDRT])

  # Data2.iat[PerfNum,40] = Data2.iloc[PerfNum,1]
    # Data2.iat[PerfNum,41] = Data2.iloc[PerfNum,6]
    # if (Data2.iloc[PerfNum,36] == '普通个寿') and (Data2.iloc[PerfNum,14] == '1年'):
        # Data2.iat[PerfNum,40] = 0.06 * Data2.iloc[PerfNum,1]


writer = pd.ExcelWriter('result.xlsx')
Data1.to_excel(writer,sheet_name = '正式保单',index = False)
Data1Pivot.to_excel(writer,sheet_name = '月度业绩汇总表')
Data2.to_excel(writer,sheet_name = '正式保单+等待回执+投保单',index = False)
Data2Pivot.to_excel(writer,sheet_name = 'MDRT')
writer.save()
# Data1Pivot = pd.read_excel('temp.xlsx')






#os.remove('temp.xlsx')
# Data1.to_excel('result.xlsx',index = False,sheet_name='sheet1')  #输出文件
# Data1Pivot.to_excel('result.xlsx',sheet_name='sheet2') 

# DataInitShape = list(DataInit.shape) #读取DataInit数据表大小并写入列表
# DataRowNum = DataInitShape[0]  #DataInit行数

#datainit = datainit[' 服务人员 '].astype('str')
#data = datainit[~datainit[9].isin(['撤单'])]
# print(datatest.info(),'\n') #数据类型
# print(datainit.dtype,'\n') #列数据类型
# print(dataimport.shape,'\n') #数据表大小
# print(datainit.shape,'\n')
# print(data.shape,'\n')

# print(datainit)


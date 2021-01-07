
import datetime
import pandas as pd
import os


# ----------待完成----------
# 计算入职月
# 计算考核期
# 入职首月业绩划分
# 当前考核期业绩汇总

print('正在处理...')

# ----------获取文件名----------
filename = os.walk('./')
xlsxname = []
for a, b, c in filename:
    xlsxname = c
# print(type(xlsxname))
listlen = len(xlsxname)

tempnum = 0
xlsxnum = 0
while tempnum < listlen:
    if ('.xlsx' in xlsxname[tempnum]) and ('公司月度业绩明细报表（个寿）' in xlsxname[tempnum]):
        xlsxnum = tempnum
    tempnum = tempnum + 1

# ----------读取文件----------
DataImport = pd.read_excel(xlsxname[xlsxnum])
DataInit = DataImport

# ----------数据预处理----------
DataInit = DataInit.astype({' 缴费方式 ': 'string', ' 首期服务人员 ': 'string', ' 服务人员工号 ': 'string', ' 保单状态 ': 'string', ' 是否医疗险 ': 'string', ' 医疗部产品分类 ': 'string', ' 分单人类型 ': 'string'})  # object转换string
DataInit[' 交单日期 '] = pd.to_datetime(DataInit[' 交单日期 '])  # object转换datetime
DataInit[' 回执操作日期 '] = pd.to_datetime(DataInit[' 回执操作日期 '])  # object转换datetime

# Data1 业绩追踪
# Data2 MDRT追踪

# Data1 选择正式保单,去掉辅助交单
Data1 = DataInit[DataInit[' 保单状态 '].str.contains('正式保单')]
Data1 = Data1.fillna({' 分单人类型 ': '非联合交单'})
Data1 = Data1.astype({' 分单人类型 ': 'string'})
Data1 = Data1[~Data1[' 分单人类型 '].str.contains('辅助交单人')]

# 选择正式保单,等待回执,投保单,去掉辅助交单
Data2 = DataInit[DataInit[' 保单状态 '].str.contains('正式保单|等待回执|投保单')]
Data2 = Data2.fillna({' 分单人类型 ': '非联合交单'})
Data2 = Data2.astype({' 分单人类型 ': 'string'})
Data2 = Data2[~Data2[' 分单人类型 '].str.contains('辅助交单人')]

# 添加业绩归属月列
DataMonPerf = pd.DataFrame(columns=['2020-01月业绩', '2020-02月业绩', '2020-03月业绩', '2020-04月业绩', '2020-05月业绩', '2020-06月业绩', '2020-07月业绩', '2020-08月业绩', '2020-09月业绩', '2020-10月业绩', '2020-11月业绩', '2020-12月业绩'])
Data1 = pd.concat([Data1, DataMonPerf])

DataMDRT = pd.DataFrame(columns=['折算保费', '2020FYC', '是否为2020业绩'])
Data2 = pd.concat([Data2, DataMDRT])

# ----------计算每月业绩----------
Data1Shape = list(Data1.shape)
Data1RowNum = Data1Shape[0]  

#收单及回执截止日
#2020
Month1912EndReceiptDate = datetime.datetime(2020, 2, 14)

Month2001StartOrderDate = datetime.datetime(2020, 1, 23)
Month2001EndOrderDate = datetime.datetime(2020, 2, 25)
Month2001EndReceiptDate = datetime.datetime(2020, 3, 10)

Month2002StartOrderDate = datetime.datetime(2020, 2, 26)
Month2002EndOrderDate = datetime.datetime(2020, 3, 25)
Month2002EndReceiptDate = datetime.datetime(2020, 4, 10)

Month2003StartOrderDate = datetime.datetime(2020, 3, 26)
Month2003EndOrderDate = datetime.datetime(2020, 4, 27)
Month2003EndReceiptDate = datetime.datetime(2020, 5, 12)

Month2004StartOrderDate = datetime.datetime(2020, 4, 28)
Month2004EndOrderDate = datetime.datetime(2020, 5, 25)
Month2004EndReceiptDate = datetime.datetime(2020, 6, 10)

Month2005StartOrderDate = datetime.datetime(2020, 5, 26)
Month2005EndOrderDate = datetime.datetime(2020, 6, 24)
Month2005EndReceiptDate = datetime.datetime(2020, 7, 10)

Month2006StartOrderDate = datetime.datetime(2020, 6, 25)
Month2006EndOrderDate = datetime.datetime(2020, 7, 27)
Month2006EndReceiptDate = datetime.datetime(2020, 8, 10)

Month2007StartOrderDate = datetime.datetime(2020, 7, 28)
Month2007EndOrderDate = datetime.datetime(2020, 8, 25)
Month2007EndReceiptDate = datetime.datetime(2020, 9, 10)

Month2008StartOrderDate = datetime.datetime(2020, 8, 26)
Month2008EndOrderDate = datetime.datetime(2020, 9, 25)
Month2008EndReceiptDate = datetime.datetime(2020, 10, 12)

Month2009StartOrderDate = datetime.datetime(2020, 9, 26)
Month2009EndOrderDate = datetime.datetime(2020, 10, 26)
Month2009EndReceiptDate = datetime.datetime(2020, 11, 12)

Month2010StartOrderDate = datetime.datetime(2020, 10, 27)
Month2010EndOrderDate = datetime.datetime(2020, 11, 25)
Month2010EndReceiptDate = datetime.datetime(2020, 12, 10)

Month2011StartOrderDate = datetime.datetime(2020, 11, 26)
Month2011EndOrderDate = datetime.datetime(2020, 12, 25)
Month2011EndReceiptDate = datetime.datetime(2021, 1, 11)

Month2012StartOrderDate = datetime.datetime(2020, 12, 26)
Month2012EndOrderDate = datetime.datetime(2021, 1, 25)
Month2012EndReceiptDate = datetime.datetime(2021, 2, 9)

#2021
Month2101StartOrderDate = datetime.datetime(2021, 1, 26)
Month2101EndOrderDate = datetime.datetime(2021, 2, 25)
Month2101EndReceiptDate = datetime.datetime(2021, 3, 10)

Month2102StartOrderDate = datetime.datetime(2021, 2, 26)
Month2102EndOrderDate = datetime.datetime(2021, 3, 25)
Month2102EndReceiptDate = datetime.datetime(2021, 4, 12)

Month2103StartOrderDate = datetime.datetime(2021, 3, 26)
Month2103EndOrderDate = datetime.datetime(2021, 4, 26)
Month2103EndReceiptDate = datetime.datetime(2021, 5, 13)

Month2104StartOrderDate = datetime.datetime(2021, 4, 27)
Month2104EndOrderDate = datetime.datetime(2021, 5, 25)
Month2104EndReceiptDate = datetime.datetime(2021, 6, 10)

Month2105StartOrderDate = datetime.datetime(2021, 5, 26)
Month2105EndOrderDate = datetime.datetime(2021, 6, 25)
Month2105EndReceiptDate = datetime.datetime(2021, 7, 12)

Month2106StartOrderDate = datetime.datetime(2021, 6, 26)
Month2106EndOrderDate = datetime.datetime(2021, 7, 26)
Month2106EndReceiptDate = datetime.datetime(2021, 8, 10)

Month2107StartOrderDate = datetime.datetime(2021, 7, 27)
Month2107EndOrderDate = datetime.datetime(2021, 8, 25)
Month2107EndReceiptDate = datetime.datetime(2021, 9, 10)

Month2108StartOrderDate = datetime.datetime(2021, 8, 26)
Month2108EndOrderDate = datetime.datetime(2021, 9, 26)
Month2108EndReceiptDate = datetime.datetime(2021, 10, 13)

Month2109StartOrderDate = datetime.datetime(2021, 9, 27)
Month2109EndOrderDate = datetime.datetime(2021, 10, 25)
Month2109EndReceiptDate = datetime.datetime(2021, 11, 10)

Month2110StartOrderDate = datetime.datetime(2021, 10, 26)
Month2110EndOrderDate = datetime.datetime(2021, 11, 25)
Month2110EndReceiptDate = datetime.datetime(2021, 12, 10)

Month2111StartOrderDate = datetime.datetime(2021, 11, 26)
Month2111EndOrderDate = datetime.datetime(2021, 12, 27)
Month2111EndReceiptDate = datetime.datetime(2021, 1, 11)

Month2112StartOrderDate = datetime.datetime(2021, 12, 28)
Month2112EndOrderDate = datetime.datetime(2022, 1, 25)
Month2112EndReceiptDate = datetime.datetime(2022, 2, 14)


# 1月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2001StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2001EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2001EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2001StartOrderDate) and (Month1912EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2001EndReceiptDate)):
        Data1.iat[PerfNum, 42] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 42] = 0
    PerfNum = PerfNum + 1

# 2月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2002StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2002EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2002EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2002StartOrderDate) and (Month2001EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2002EndReceiptDate)):
        Data1.iat[PerfNum, 43] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 43] = 0
    PerfNum = PerfNum + 1

# 3月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2003StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2003EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2003EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2003StartOrderDate) and (Month2002EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2003EndReceiptDate)):
        Data1.iat[PerfNum, 44] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 44] = 0
    PerfNum = PerfNum + 1

# 4月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2004StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2004EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2004EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2004StartOrderDate) and (Month2003EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2004EndReceiptDate)):
        Data1.iat[PerfNum, 45] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 45] = 0
    PerfNum = PerfNum + 1

# 5月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2005StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2005EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2005EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2005StartOrderDate) and (Month2004EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2005EndReceiptDate)):
        Data1.iat[PerfNum, 46] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 46] = 0
    PerfNum = PerfNum + 1

# 6月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2006StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2006EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2006EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2006StartOrderDate) and (Month2005EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2006EndReceiptDate)):
        Data1.iat[PerfNum, 47] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 47] = 0
    PerfNum = PerfNum + 1

# 7月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2007StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2007EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2007EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2007StartOrderDate) and (Month2006EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2007EndReceiptDate)):
        Data1.iat[PerfNum, 48] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 48] = 0
    PerfNum = PerfNum + 1

# 8月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2008StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2008EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2008EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2008StartOrderDate) and (Month2007EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2008EndReceiptDate)):
        Data1.iat[PerfNum, 49] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 49] = 0
    PerfNum = PerfNum + 1

# 9月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2009StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2009EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2009EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2009StartOrderDate) and (Month2008EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2009EndReceiptDate)):
        Data1.iat[PerfNum, 50] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 50] = 0
    PerfNum = PerfNum + 1

# 10月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2010StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2010EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2010EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2010StartOrderDate) and (Month2009EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2010EndReceiptDate)):
        Data1.iat[PerfNum, 51] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 51] = 0
    PerfNum = PerfNum + 1

# 11月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2011StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2011EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2011EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2011StartOrderDate) and (Month2010EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2011EndReceiptDate)):
        Data1.iat[PerfNum, 52] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 52] = 0
    PerfNum = PerfNum + 1

# 12月
PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2012StartOrderDate <= Data1.iloc[PerfNum, 25]) and (Data1.iloc[PerfNum, 25] <= Month2012EndOrderDate) and (Data1.iloc[PerfNum, 28] <= Month2012EndReceiptDate)) or ((Data1.iloc[PerfNum, 25] < Month2012StartOrderDate) and (Month2011EndReceiptDate < Data1.iloc[PerfNum, 28]) and (Data1.iloc[PerfNum, 28] <= Month2012EndReceiptDate)):
        Data1.iat[PerfNum, 53] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 53] = 0
    PerfNum = PerfNum + 1

#月度业绩汇总表
Data1Pivot = pd.pivot_table(Data1, values=['2020-01月业绩', '2020-02月业绩', '2020-03月业绩', '2020-04月业绩', '2020-05月业绩', '2020-06月业绩', '2020-07月业绩', '2020-08月业绩', '2020-09月业绩', '2020-10月业绩', '2020-11月业绩', '2020-12月业绩'], index=' 首期服务人员 ', aggfunc='sum')


# 计算MDRT折算保费,FYC
Data2Shape = list(Data2.shape)
Data2RowNum = Data2Shape[0]

#筛选2020业绩
PerfNum = 0
while PerfNum < Data2RowNum:
    if (((Month2001StartOrderDate <= Data2.iloc[PerfNum, 25]) and (Data2.iloc[PerfNum, 25] <= Month2012EndOrderDate) and (Data2.iloc[PerfNum, 28] <= Month2012EndReceiptDate)) or ((Data2.iloc[PerfNum, 25] < Month2001StartOrderDate) and (Month1912EndReceiptDate < Data2.iloc[PerfNum, 28]) and (Data2.iloc[PerfNum, 28] <= Month2012EndReceiptDate))):
        Data2.iat[PerfNum, 44] = 1
    else:
        Data2.iat[PerfNum, 44] = 0
    if ((Data2.iloc[PerfNum, 29] == '等待回执') or (Data2.iloc[PerfNum, 29] == '投保单')):
        Data2.iat[PerfNum, 44] = 1
    PerfNum = PerfNum + 1

#计算折算保费
PerfNum = 0
while PerfNum < Data2RowNum:
    if Data2.iloc[PerfNum, 44] == 1:
        Data2.iat[PerfNum, 42] = Data2.iloc[PerfNum, 1]
        Data2.iat[PerfNum, 43] = Data2.iloc[PerfNum, 6]
    else:
        Data2.iat[PerfNum, 42] = 0
        Data2.iat[PerfNum, 43] = 0
    if (Data2.iloc[PerfNum, 44] == 1 and (Data2.iloc[PerfNum, 38] == '普通个寿')) and (Data2.iloc[PerfNum, 14] == '1年'):
        Data2.iat[PerfNum, 42] = 0.06 * Data2.iloc[PerfNum, 1]
    PerfNum = PerfNum + 1

#MDRT数据汇总
Data2Pivot = pd.pivot_table(Data2, values=['折算保费', '2020FYC'], index=' 首期服务人员 ', aggfunc='sum')
Data2Pivot = Data2Pivot.sort_values(by=["折算保费"], ascending=False)
Data2MDRT = pd.DataFrame(columns=['2020-MDRT-FYC差值', '2020-MDRT-保费差值','2020-COT-FYC差值', '2020-COT-保费差值','2020-TOT-FYC差值', '2020-TOT-保费差值'])
Data2Pivot = pd.concat([Data2Pivot,Data2MDRT])

Data2PivotShape = list(Data2Pivot.shape)
Data2PivotRowNum = Data2PivotShape[0]

#MDRT标准

MDRTCommission20 = 171300
MDRTPremium20 = 513900
COTCommission20 = 3 * MDRTCommission20
COTPremium20 = 3 * MDRTPremium20
TOTCommission20 = 6 * MDRTCommission20
TOTPremium20 = 6 * MDRTPremium20

MDRTCommission20 = 183100
MDRTPremium20 = 549300
COTCommission20 = 3 * MDRTCommission20
COTPremium20 = 3 * MDRTPremium20
TOTCommission20 = 6 * MDRTCommission20
TOTPremium20 = 6 * MDRTPremium20


PerfNum = 0
while PerfNum < Data2PivotRowNum:
    if Data2Pivot.iloc[PerfNum,0] >= MDRTCommission20:
        Data2Pivot.iat[PerfNum,2] = '预达成'
    else:
        Data2Pivot.iat[PerfNum,2] = MDRTCommission20 - Data2Pivot.iloc[PerfNum,0]
    PerfNum = PerfNum + 1

PerfNum = 0
while PerfNum < Data2PivotRowNum:
    if Data2Pivot.iloc[PerfNum,1] >= MDRTPremium20:
        Data2Pivot.iat[PerfNum,3] = '预达成'
    else:
        Data2Pivot.iat[PerfNum,3] = MDRTPremium20 - Data2Pivot.iloc[PerfNum,1]
    PerfNum = PerfNum + 1

PerfNum = 0
while PerfNum < Data2PivotRowNum:
    if Data2Pivot.iloc[PerfNum,0] >= COTCommission20:
        Data2Pivot.iat[PerfNum,4] = '预达成'
    else:
        Data2Pivot.iat[PerfNum,4] = COTCommission20 - Data2Pivot.iloc[PerfNum,0]
    PerfNum = PerfNum + 1

PerfNum = 0
while PerfNum < Data2PivotRowNum:
    if Data2Pivot.iloc[PerfNum,1] >= COTPremium20:
        Data2Pivot.iat[PerfNum,5] = '预达成'
    else:
        Data2Pivot.iat[PerfNum,5] = COTPremium20 - Data2Pivot.iloc[PerfNum,1]
    PerfNum = PerfNum + 1

PerfNum = 0
while PerfNum < Data2PivotRowNum:
    if Data2Pivot.iloc[PerfNum,0] >= TOTCommission20:
        Data2Pivot.iat[PerfNum,6] = '预达成'
    else:
        Data2Pivot.iat[PerfNum,6] = TOTCommission20 - Data2Pivot.iloc[PerfNum,0]
    PerfNum = PerfNum + 1

PerfNum = 0
while PerfNum < Data2PivotRowNum:
    if Data2Pivot.iloc[PerfNum,1] >= TOTPremium20:
        Data2Pivot.iat[PerfNum,7] = '预达成'
    else:
        Data2Pivot.iat[PerfNum,7] = TOTPremium20 - Data2Pivot.iloc[PerfNum,1]
    PerfNum = PerfNum + 1

#----------数据写入xlsx----------
writer = pd.ExcelWriter('result.xlsx')
Data1.to_excel(writer, sheet_name = '正式保单', index = False)
Data1Pivot.to_excel(writer, sheet_name = '月度业绩汇总表')
Data2.to_excel(writer, sheet_name = '正式保单+等待回执+投保单', index = False)
Data2Pivot.to_excel(writer, sheet_name = 'MDRT')
writer.save()

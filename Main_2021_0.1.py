
import datetime
import pandas as pd
import os


# ----------待完成----------
# 2021年月业绩计算逻辑
# 2021年月业绩表头
# 2021年MDRT计算逻辑
# 2021年MDRT表头
# 
# 
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
DataInit = DataInit.astype({' 缴费方式 ': 'string', ' 服务人员 ': 'string', ' 服务人员工号 ': 'string', ' 保单状态 ': 'string', ' 是否医疗险 ': 'string', ' 医疗部产品分类 ': 'string', ' 分单人类型 ': 'string'})  # object转换string
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
DataMonPerf = pd.DataFrame(columns=['2021-01月业绩', '2021-02月业绩', '2021-03月业绩', '2021-04月业绩', '2021-05月业绩', '2021-06月业绩', '2021-07月业绩', '2021-08月业绩', '2021-09月业绩', '2021-10月业绩', '2021-11月业绩', '2021-12月业绩'])
Data1 = pd.concat([Data1, DataMonPerf])

DataMDRT = pd.DataFrame(columns=['折算保费', '2021FYC', '是否为2021业绩'])
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




PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2101StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2101EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2101EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2101StartOrderDate) and (Month2012EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2101EndReceiptDate)):
        Data1.iat[PerfNum, 40] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 40] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2102StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2102EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2102EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2102StartOrderDate) and (Month2101EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2102EndReceiptDate)):
        Data1.iat[PerfNum, 41] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 41] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2103StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2103EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2103EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2103StartOrderDate) and (Month2102EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2103EndReceiptDate)):
        Data1.iat[PerfNum, 42] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 42] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2104StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2104EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2104EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2104StartOrderDate) and (Month2103EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2104EndReceiptDate)):
        Data1.iat[PerfNum, 43] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 43] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2105StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2105EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2105EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2105StartOrderDate) and (Month2104EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2105EndReceiptDate)):
        Data1.iat[PerfNum, 44] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 44] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2106StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2106EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2106EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2106StartOrderDate) and (Month2105EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2106EndReceiptDate)):
        Data1.iat[PerfNum, 45] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 45] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2107StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2107EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2107EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2107StartOrderDate) and (Month2106EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2107EndReceiptDate)):
        Data1.iat[PerfNum, 46] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 46] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2108StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2108EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2108EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2108StartOrderDate) and (Month2107EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2108EndReceiptDate)):
        Data1.iat[PerfNum, 47] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 47] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2109StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2109EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2109EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2109StartOrderDate) and (Month2108EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2109EndReceiptDate)):
        Data1.iat[PerfNum, 48] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 48] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2110StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2110EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2110EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2110StartOrderDate) and (Month2109EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2110EndReceiptDate)):
        Data1.iat[PerfNum, 49] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 49] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2111StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2111EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2111EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2111StartOrderDate) and (Month2110EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2111EndReceiptDate)):
        Data1.iat[PerfNum, 50] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 50] = 0
    PerfNum = PerfNum + 1


PerfNum = 0
while PerfNum < Data1RowNum:
    if ((Month2112StartOrderDate <= Data1.iloc[PerfNum, 23]) and (Data1.iloc[PerfNum, 23] <= Month2112EndOrderDate) and (Data1.iloc[PerfNum, 26] <= Month2112EndReceiptDate)) or ((Data1.iloc[PerfNum, 23] < Month2112StartOrderDate) and (Month2111EndReceiptDate < Data1.iloc[PerfNum, 26]) and (Data1.iloc[PerfNum, 26] <= Month2112EndReceiptDate)):
        Data1.iat[PerfNum, 51] = Data1.iloc[PerfNum, 7]
    else:
        Data1.iat[PerfNum, 51] = 0
    PerfNum = PerfNum + 1


#月度业绩汇总表
Data1Pivot = pd.pivot_table(Data1, values=['2021-01月业绩', '2021-02月业绩', '2021-03月业绩', '2021-04月业绩', '2021-05月业绩', '2021-06月业绩', '2021-07月业绩', '2021-08月业绩', '2021-09月业绩', '2021-10月业绩', '2021-11月业绩', '2021-12月业绩'], index=' 服务人员 ', aggfunc='sum')


# 计算MDRT折算保费,FYC
Data2Shape = list(Data2.shape)
Data2RowNum = Data2Shape[0]

#筛选2021业绩
PerfNum = 0
while PerfNum < Data2RowNum:
    if (((Month2101StartOrderDate <= Data2.iloc[PerfNum, 23]) and (Data2.iloc[PerfNum, 23] <= Month2112EndOrderDate) and (Data2.iloc[PerfNum, 26] <= Month2112EndReceiptDate)) or ((Data2.iloc[PerfNum, 23] < Month2101StartOrderDate) and (Month2012EndReceiptDate < Data2.iloc[PerfNum, 26]) and (Data2.iloc[PerfNum, 26] <= Month2112EndReceiptDate))):
        Data2.iat[PerfNum, 42] = 1
    else:
        Data2.iat[PerfNum, 42] = 0
    if ((Data2.iloc[PerfNum, 27] == '等待回执') or (Data2.iloc[PerfNum, 27] == '投保单')) and (datetime.datetime.now() > Month2012EndReceiptDate):
        Data2.iat[PerfNum, 42] = 1
    PerfNum = PerfNum + 1


#计算折算保费
PerfNum = 0
while PerfNum < Data2RowNum:
    if Data2.iloc[PerfNum, 42] == 1:
        Data2.iat[PerfNum, 40] = Data2.iloc[PerfNum, 1]
        Data2.iat[PerfNum, 41] = Data2.iloc[PerfNum, 6]
    else:
        Data2.iat[PerfNum, 40] = 0
        Data2.iat[PerfNum, 41] = 0
    if (Data2.iloc[PerfNum, 42] == 1 and (Data2.iloc[PerfNum, 36] == '普通个寿')) and (Data2.iloc[PerfNum, 14] == '1年'):
        Data2.iat[PerfNum, 40] = 0.06 * Data2.iloc[PerfNum, 1]
    PerfNum = PerfNum + 1

#MDRT数据汇总
Data2Pivot = pd.pivot_table(Data2, values=['折算保费', '2021FYC'], index=' 服务人员 ', aggfunc='sum')
Data2Pivot = Data2Pivot.sort_values(by=["折算保费"], ascending=False)
Data2MDRT = pd.DataFrame(columns=['2021-MDRT-FYC差值', '2021-MDRT-保费差值','2021-COT-FYC差值', '2021-COT-保费差值','2021-TOT-FYC差值', '2020-TOT-保费差值'])
Data2Pivot = pd.concat([Data2Pivot,Data2MDRT])

Data2PivotShape = list(Data2Pivot.shape)
Data2PivotRowNum = Data2PivotShape[0]

#MDRT标准
MDRTCommission = 171300
MDRTPremium = 513900
COTCommission = 3 * MDRTCommission
COTPremium = 3 * MDRTPremium
TOTCommission = 6 * MDRTCommission
TOTPremium = 6 * MDRTPremium

# PerfNum = 0
# while PerfNum < Data2PivotRowNum:
#     if Data2Pivot.iloc[PerfNum,0] >= MDRTCommission:
#         Data2Pivot.iat[PerfNum,2] = '预达成'
#     else:
#         Data2Pivot.iat[PerfNum,2] = MDRTCommission - Data2Pivot.iloc[PerfNum,0]
#     PerfNum = PerfNum + 1

# PerfNum = 0
# while PerfNum < Data2PivotRowNum:
#     if Data2Pivot.iloc[PerfNum,1] >= MDRTPremium:
#         Data2Pivot.iat[PerfNum,3] = '预达成'
#     else:
#         Data2Pivot.iat[PerfNum,3] = MDRTPremium - Data2Pivot.iloc[PerfNum,1]
#     PerfNum = PerfNum + 1

# PerfNum = 0
# while PerfNum < Data2PivotRowNum:
#     if Data2Pivot.iloc[PerfNum,0] >= COTCommission:
#         Data2Pivot.iat[PerfNum,4] = '预达成'
#     else:
#         Data2Pivot.iat[PerfNum,4] = COTCommission - Data2Pivot.iloc[PerfNum,0]
#     PerfNum = PerfNum + 1

# PerfNum = 0
# while PerfNum < Data2PivotRowNum:
#     if Data2Pivot.iloc[PerfNum,1] >= COTPremium:
#         Data2Pivot.iat[PerfNum,5] = '预达成'
#     else:
#         Data2Pivot.iat[PerfNum,5] = COTPremium - Data2Pivot.iloc[PerfNum,1]
#     PerfNum = PerfNum + 1

# PerfNum = 0
# while PerfNum < Data2PivotRowNum:
#     if Data2Pivot.iloc[PerfNum,0] >= TOTCommission:
#         Data2Pivot.iat[PerfNum,6] = '预达成'
#     else:
#         Data2Pivot.iat[PerfNum,6] = TOTCommission - Data2Pivot.iloc[PerfNum,0]
#     PerfNum = PerfNum + 1

# PerfNum = 0
# while PerfNum < Data2PivotRowNum:
#     if Data2Pivot.iloc[PerfNum,1] >= TOTPremium:
#         Data2Pivot.iat[PerfNum,7] = '预达成'
#     else:
#         Data2Pivot.iat[PerfNum,7] = TOTPremium - Data2Pivot.iloc[PerfNum,1]
#     PerfNum = PerfNum + 1

#----------数据写入xlsx----------
writer = pd.ExcelWriter('result.xlsx')
Data1.to_excel(writer, sheet_name = '正式保单', index = False)
Data1Pivot.to_excel(writer, sheet_name = '月度业绩汇总表')
Data2.to_excel(writer, sheet_name = '正式保单+等待回执+投保单', index = False)
Data2Pivot.to_excel(writer, sheet_name = 'MDRT')
writer.save()

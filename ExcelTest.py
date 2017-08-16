import xlrd
import xlwt
import sys

# 当月（2017年几月）
currentMonth = 6
#  路径前加r，原因：文件名中的 \U 开始的字符被编译器认为是八进制
#  保存输出数据的文档地址
file_path = r"C:\Users\Zhang Yu\Desktop\数据结果.xls"
#  原文档地址
sourcefile_path = r"C:\Users\Zhang Yu\Desktop\微机站订单（基于分类+协议）.xlsx"


data = xlrd.open_workbook(sourcefile_path)
booksheet = data.sheets()[0]  # 打开第一张表
wb = xlwt.Workbook(encoding='utf-8', style_compression=0)
newsheet = wb.add_sheet('datasheet', cell_overwrite_ok=True)  # 创建sheet2 第二参数用于确认同一个cell单元是否可以重设值。

# 找到订单编号列
colPO = -1
# 找到下单时间列
colPlaceTime = -1
# 找到订单状态列
colOrderStatus = -1
# 找到实际采购数量列
colQuantity = -1
# 找到单个净价列
colTotalPrice = -1
# 找到省份名称列
colProvince = -1
# 找到分类列
colCategory = -1
# 找到供应商列
colSupplier = -1
# 找到收货时间列
colArrivalTime = -1

# python3.0后 xrange和range一样
for col in range(booksheet.ncols):
    cell = booksheet.cell_value(0, col)
    if cell == "订单编号":
        colPO = col
    if cell == "下单时间":
        colPlaceTime = col
    if cell == "订单状态":
        colOrderStatus = col
    if cell == "实际采购数量":
        colQuantity = col
    if cell == "单个净价":
        colTotalPrice = col
    if cell == "省份名称":
        colProvince = col
    if cell == "分类":
        colCategory = col
    if cell == "供应商":
        colSupplier = col
    if cell == "收货时间":
        colArrivalTime = col
# 若某一列名没有找到，则报错
if colPO == -1:
    print("没有找到订单编号列！")
    sys.exit(0)
if colPlaceTime == -1:
    print("没有找到下单时间列！")
    sys.exit(0)
if colOrderStatus == -1:
    print("没有找到订单状态列！")
    sys.exit(0)
if colQuantity == -1:
    print("没有找到实际采购数量列！")
    sys.exit(0)
if colTotalPrice == -1:
    print("没有找到单个净价列！")
    sys.exit(0)
if colProvince == -1:
    print("没有找到省份名称列！")
    sys.exit(0)
if colCategory == -1:
    print("没有找到分类列！")
    sys.exit(0)
if colSupplier == -1:
    print("没有找到供应商列！")
    sys.exit(0)
if colArrivalTime == -1:
    print("没有找到收货时间列！")
    sys.exit(0)

# 2016与2017年订单集合
POOfY = {2015:set(),2016:set(),2017:set()}
# index从0-11分别代表1月到12月  各月净价总和
ListOfTotalPrice2015 = [0,0,0,0,0,0,0,0,0,0,0,0]
ListOfTotalPrice2016 = [0,0,0,0,0,0,0,0,0,0,0,0]
ListOfTotalPrice2017 = [0,0,0,0,0,0,0,0,0,0,0,0]
# 各年中各个月对应订单集合
POOf2015 = {1:set(),2:set(),3:set(),4:set(),5:set(),6:set(),7:set(),8:set(),9:set(),10:set(),11:set(),12:set()}
POOf2016 = {1:set(),2:set(),3:set(),4:set(),5:set(),6:set(),7:set(),8:set(),9:set(),10:set(),11:set(),12:set()}
POOf2017 = {1:set(),2:set(),3:set(),4:set(),5:set(),6:set(),7:set(),8:set(),9:set(),10:set(),11:set(),12:set()}

# 以下字典key均为省分名称
# 2017年省分与订单集合的字典
ProAndPO = {}
# 2017年省分与交易额的字典
ProAndTotalPrice = {}
# 特殊月份省分与订单集合的字典
SeProAndPO = {}
# 特殊月份省分与交易额的字典
SeProAndTotalPrice = {}
# 累计总省分与订单集合的字典
TotalProAndPO = {}
# 累计总省分与交易额的字典
TotalProAndTotalPrice = {}
# 典配—2017年省分与订单集合的字典
CaProAndPO = {}
# 典配—特殊月份省分与订单集合的字典
CaSeProAndPO = {}
# 典配—累计总省分与订单集合的字典
CaTotalProAndPO = {}
# 本省2017年的供应商
SupplierOfPro = {}
# 本省2017年特殊月的供应商
SeSupplierOfPro = {}
# 本省2017年的供应商
TotalSupplierOfPro = {}
# 2017年已到货到货订单省分与其订单天数list字典
ProvinceAndArrivalDay = {}
# 2017年特殊月已到货到货订单省分与其订单天数list字典
SeProvinceAndArrivalDay = {}

# 以下字典key均为供应商名称
# 2017年供应商与订单集合字典
POofSupplier = {}
# 2017年供应商与省分集合字典
ProvinceAndSupplier = {}
# 2017年特殊月供应商与订单集合字典
SePOofSupplier = {}
# 2017年特殊月供应商与省分集合字典
SeProvinceAndSupplier = {}
# 2017年供应商与订单集合字典
TotalPOofSupplier = {}
# 2017年供应商与省分集合字典
TotalProvinceAndSupplier = {}


# 按年维度处理每个月总净价和订单list
def yearlyCount(pricelist,polist):
    monthlyCount(1, pricelist, polist)
    monthlyCount(2, pricelist, polist)
    monthlyCount(3, pricelist, polist)
    monthlyCount(4, pricelist, polist)
    monthlyCount(5, pricelist, polist)
    monthlyCount(6, pricelist, polist)
    monthlyCount(7, pricelist, polist)
    monthlyCount(8, pricelist, polist)
    monthlyCount(9, pricelist, polist)
    monthlyCount(10, pricelist, polist)
    monthlyCount(11, pricelist, polist)
    monthlyCount(12, pricelist, polist)

# 按月维度处理每个月总净价和订单list
def monthlyCount(month,pricelist,polist):
    if date_value[1] == month:
        pricelist[month-1] += booksheet.cell_value(row, colTotalPrice) * booksheet.cell_value(row, colQuantity)
        polist[month].add(booksheet.cell_value(row, colPO))

# 供应商维度统计，包括（供应商-订单）（供应商-省分）
def supplierCount(poofsupplier,provinceofsupplier):
    if (booksheet.cell_value(row, colSupplier) not in poofsupplier):
        temp = set()
        temp.add(booksheet.cell_value(row, colPO))
        poofsupplier[booksheet.cell_value(row, colSupplier)] = temp
        stemp = set()
        stemp.add(booksheet.cell_value(row, colProvince)[0:2])
        provinceofsupplier[booksheet.cell_value(row, colSupplier)] = stemp
    else:
        temp = poofsupplier[booksheet.cell_value(row, colSupplier)]
        temp.add(booksheet.cell_value(row, colPO))
        poofsupplier[booksheet.cell_value(row, colSupplier)] = temp
        stemp = provinceofsupplier[booksheet.cell_value(row, colSupplier)]
        stemp.add(booksheet.cell_value(row, colProvince)[0:2])
        provinceofsupplier[booksheet.cell_value(row, colSupplier)] = stemp

# 典配或组合商品统计
def combinationCount(caproandpo):
    if ("组合商品" in booksheet.cell_value(row, colCategory) or "典配" in booksheet.cell_value(row, colCategory)):
        if (booksheet.cell_value(row, colProvince)[0:2] not in caproandpo):
            temp = set()
            temp.add(booksheet.cell_value(row, colPO))
            caproandpo[booksheet.cell_value(row, colProvince)[0:2]] = temp
        else:
            temp = caproandpo[booksheet.cell_value(row, colProvince)[0:2]]
            temp.add(booksheet.cell_value(row, colPO))
            caproandpo[booksheet.cell_value(row, colProvince)[0:2]] = temp

# 省分订单与总价统计
def orderAndPriceCountByPro(proandpo,proandtotalprice):
    if (booksheet.cell_value(row, colProvince)[0:2] not in proandpo):
        temp = set()
        temp.add(booksheet.cell_value(row, colPO))
        proandpo[booksheet.cell_value(row, colProvince)[0:2]] = temp
        proandtotalprice[booksheet.cell_value(row, colProvince)[0:2]] = \
            booksheet.cell_value(row, colTotalPrice) * booksheet.cell_value(row, colQuantity)
    else:
        temp = proandpo[booksheet.cell_value(row, colProvince)[0:2]]
        temp.add(booksheet.cell_value(row, colPO))
        proandpo[booksheet.cell_value(row, colProvince)[0:2]] = temp
        proandtotalprice[booksheet.cell_value(row, colProvince)[0:2]] += \
            booksheet.cell_value(row, colTotalPrice) * booksheet.cell_value(row, colQuantity)

# 由（供应商-省分）字典转化为（省分-供应商）字典
def dictTransform(provinceandsupplier,supplierofpro):
    for key in provinceandsupplier:
        for p in provinceandsupplier[key]:
            if p not in supplierofpro:
                temp = set()
                temp.add(key)
                supplierofpro[p] = temp
            else:
                temp = supplierofpro[p]
                temp.add(key)
                supplierofpro[p] = temp


for row in range(booksheet.nrows):
    if row == 0:  # 跳过第一行
        continue
    if (booksheet.cell_value(row, colOrderStatus)=="确认订单" or booksheet.cell_value(row, colOrderStatus)=="完成订单"):
        date_value = xlrd.xldate_as_tuple(booksheet.cell_value(row, colPlaceTime), data.datemode)
        # 2015年统计
        if date_value[0] == 2015:
            POOfY[2015].add(booksheet.cell_value(row, colPO))
            yearlyCount(ListOfTotalPrice2015, POOf2015)
        # 2016年统计
        if date_value[0] == 2016:
            POOfY[2016].add(booksheet.cell_value(row, colPO))
            yearlyCount(ListOfTotalPrice2016,POOf2016)
        # 2017年统计
        if date_value[0] == 2017:
            POOfY[2017].add(booksheet.cell_value(row, colPO))
            yearlyCount(ListOfTotalPrice2017, POOf2017)

            # 2017年后统计
            if (date_value[1] <= currentMonth):

                # 省分订单与总价统计
                orderAndPriceCountByPro(ProAndPO, ProAndTotalPrice)

                # 统计典配
                combinationCount(CaProAndPO)

                # 2017年供应商维度统计
                supplierCount(POofSupplier, ProvinceAndSupplier)

            # 特殊月为当月 特殊月统计
            if (date_value[1]==currentMonth):

                # 省分订单与总价统计
                orderAndPriceCountByPro(SeProAndPO, SeProAndTotalPrice)

                # 典配统计
                combinationCount(CaSeProAndPO)

                # 供应商维度统计
                supplierCount(SePOofSupplier, SeProvinceAndSupplier)

        # 无论年份
        # 省分订单与总价统计
        orderAndPriceCountByPro(TotalProAndPO, TotalProAndTotalPrice)

        # 典配统计
        combinationCount(CaTotalProAndPO)

        # 供应商维度统计
        supplierCount(TotalPOofSupplier, TotalProvinceAndSupplier)

        # 到货天数统计（统计的为到货时间为2017年后的）
        # 若到货时间列不为空
        if booksheet.cell_value(row, colArrivalTime):
            date_value_arrival = xlrd.xldate_as_tuple(booksheet.cell_value(row, colArrivalTime), data.datemode)
            # 统计的为到货时间为2017年后的
            if date_value_arrival[0] == 2017:
                if (booksheet.cell_value(row, colProvince)[0:2] not in ProvinceAndArrivalDay):
                    templist = []
                    templist.append(round(booksheet.cell_value(row, colArrivalTime)-booksheet.cell_value(row, colPlaceTime)))
                    ProvinceAndArrivalDay[booksheet.cell_value(row, colProvince)[0:2]] = templist
                else:
                    templist = ProvinceAndArrivalDay[booksheet.cell_value(row, colProvince)[0:2]]
                    templist.append(round(booksheet.cell_value(row, colArrivalTime)-booksheet.cell_value(row, colPlaceTime)))
                    ProvinceAndArrivalDay[booksheet.cell_value(row, colProvince)[0:2]] = templist
                # 统计的为到货时间为指定月的
                if date_value_arrival[1] == currentMonth:
                    if (booksheet.cell_value(row, colProvince)[0:2] not in SeProvinceAndArrivalDay):
                        templist = []
                        templist.append(
                            round(booksheet.cell_value(row, colArrivalTime) - booksheet.cell_value(row, colPlaceTime)))
                        SeProvinceAndArrivalDay[booksheet.cell_value(row, colProvince)[0:2]] = templist
                    else:
                        templist = SeProvinceAndArrivalDay[booksheet.cell_value(row, colProvince)[0:2]]
                        templist.append(
                            round(booksheet.cell_value(row, colArrivalTime) - booksheet.cell_value(row, colPlaceTime)))
                        SeProvinceAndArrivalDay[booksheet.cell_value(row, colProvince)[0:2]] = templist
# 由（供应商-省分）字典转化为（省分-供应商）字典
dictTransform(ProvinceAndSupplier,SupplierOfPro)
dictTransform(SeProvinceAndSupplier,SeSupplierOfPro)
dictTransform(TotalProvinceAndSupplier,TotalSupplierOfPro)


#  扩大1到11列的宽度
for n in range(0, 11):
    newsheet.col(n).width = 256*20

#  表头格式
header_style = xlwt.easyxf('font:height 540;')
#  表行列名格式
tablestyle = 'pattern: pattern solid, fore_colour yellow; '  # 背景颜色为黄色
tablestyle += 'font: height 200, bold on; '  #  粗体字
tablestyle += 'align: horz centre, vert center; '  #  居中
table_style = xlwt.easyxf(tablestyle)
#  正文格式
textstyle = 'font: height 200;'  #  粗体字
textstyle += 'align: horz centre, vert center; '  #  居中
text_style = xlwt.easyxf(textstyle)


x = 0
y = 0
newsheet.write(x+0, y+0, '累计交易情况', header_style)

newsheet.write(x+1, y+0, '年份', table_style)
newsheet.write(x+2, y+0, '订单量', table_style)

newsheet.write(x+1, y+1, '2015', table_style)
newsheet.write(x+2, y+1, len(POOfY[2015]), text_style)
newsheet.write(x+1, y+2, '2016', table_style)
newsheet.write(x+2, y+2, len(POOfY[2016]), text_style)
newsheet.write(x+1, y+3, '2017', table_style)
newsheet.write(x+2, y+3, len(POOfY[2017]), text_style)

newsheet.write(x+1, y+4, '总计', table_style)
newsheet.write(x+2, y+4, len(POOfY[2015])+len(POOfY[2016])+len(POOfY[2017]), text_style)
x += 3  #  标题与内容共三行

x += 3
y = 0
newsheet.write(x+0, y+0, '交易情况分析', header_style)
newsheet.write(x+1, y+0, '年份', table_style)
newsheet.write(x+1, y+1, '月份', table_style)
newsheet.write(x+1, y+2, '订单量', table_style)
newsheet.write(x+1, y+3, '交易额', table_style)
x += 2   #  标头两行
newsheet.write_merge(x, x+11, y, y, '2015', table_style)
for n in range(0, 12):
    newsheet.write(x + n, y + 1, n + 1, table_style)
    newsheet.write(x + n, y + 2, len(POOf2015[n+1]), text_style)
    newsheet.write(x + n, y + 3, ListOfTotalPrice2015[n], text_style)
x += 12
newsheet.write_merge(x, x+11, y, y, '2016', table_style)
for n in range(0, 12):
    newsheet.write(x + n, y + 1, n + 1, table_style)
    newsheet.write(x + n, y + 2, len(POOf2016[n+1]), text_style)
    newsheet.write(x + n, y + 3, ListOfTotalPrice2016[n], text_style)
x += 12
newsheet.write_merge(x, x+11, y, y, '2017', table_style)
for n in range(0, 12):
    newsheet.write(x + n, y + 1, n + 1, table_style)
    newsheet.write(x + n, y + 2, len(POOf2017[n+1]), text_style)
    newsheet.write(x + n, y + 3, ListOfTotalPrice2017[n], text_style)
x += 12

x += 3
y = 0
newsheet.write(x+0, y+0, '省分累计下单情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '订单量', table_style)
newsheet.write(x+1, y+2, '交易额', table_style)
x += 2
for key in TotalProAndPO:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(TotalProAndPO[key]), text_style)
    newsheet.write(x, y+2, TotalProAndTotalPrice[key], text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年省分下单情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '订单量', table_style)
newsheet.write(x+1, y+2, '交易额', table_style)
x += 2
for key in ProAndPO:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(ProAndPO[key]), text_style)
    newsheet.write(x, y+2, ProAndTotalPrice[key], text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年%d' % currentMonth + '月省分下单情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '订单量', table_style)
newsheet.write(x+1, y+2, '交易额', table_style)
x += 2
for key in SeProAndPO:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(SeProAndPO[key]), text_style)
    newsheet.write(x, y+2, SeProAndTotalPrice[key], text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '省分累计典配模式下单情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '典配订单数量', table_style)
newsheet.write(x+1, y+2, '订单量', table_style)
newsheet.write(x+1, y+3, '百分比', table_style)
x += 2
for key in TotalProAndPO:
    newsheet.write(x, y, key, text_style)
    if key in CaTotalProAndPO:
        newsheet.write(x, y + 1, len(CaTotalProAndPO[key]), text_style)
        newsheet.write(x, y+2, len(TotalProAndPO[key]), text_style)
        newsheet.write(x, y+3, len(CaTotalProAndPO[key])/len(TotalProAndPO[key]), text_style)
    else:
        newsheet.write(x, y + 1, 0, text_style)
        newsheet.write(x, y+2, len(TotalProAndPO[key]), text_style)
        newsheet.write(x, y+3, 0, text_style)
    x += 1


x += 3
y = 0
newsheet.write(x+0, y+0, '2017年省分累计典配模式下单情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '典配订单数量', table_style)
newsheet.write(x+1, y+2, '订单量', table_style)
newsheet.write(x+1, y+3, '百分比', table_style)
x += 2
for key in ProAndPO:
    newsheet.write(x, y, key, text_style)
    if key in CaProAndPO:
        newsheet.write(x, y + 1, len(CaProAndPO[key]), text_style)
        newsheet.write(x, y+2, len(ProAndPO[key]), text_style)
        newsheet.write(x, y+3, len(CaProAndPO[key])/len(ProAndPO[key]), text_style)
    else:
        newsheet.write(x, y + 1, 0, text_style)
        newsheet.write(x, y+2, len(ProAndPO[key]), text_style)
        newsheet.write(x, y+3, 0, text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年%d' % currentMonth + '月省分累计典配模式下单情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '典配订单数量', table_style)
newsheet.write(x+1, y+2, '订单量', table_style)
newsheet.write(x+1, y+3, '百分比', table_style)
x += 2
for key in SeProAndPO:
    newsheet.write(x, y, key, text_style)
    if key in CaSeProAndPO:
        newsheet.write(x, y + 1, len(CaSeProAndPO[key]), text_style)
        newsheet.write(x, y+2, len(SeProAndPO[key]), text_style)
        newsheet.write(x, y+3, len(CaSeProAndPO[key])/len(SeProAndPO[key]), text_style)
    else:
        newsheet.write(x, y + 1, 0, text_style)
        newsheet.write(x, y+2, len(SeProAndPO[key]), text_style)
        newsheet.write(x, y+3, 0, text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '累计供应商份额统计', header_style)
newsheet.write(x+1, y+0, '供应商名称', table_style)
newsheet.write(x+1, y+1, '订单数量', table_style)
x += 2
for key in TotalPOofSupplier:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(TotalPOofSupplier[key]), text_style)
    x += 1


x += 3
y = 0
newsheet.write(x+0, y+0, '2017年供应商份额统计', header_style)
newsheet.write(x+1, y+0, '供应商名称', table_style)
newsheet.write(x+1, y+1, '订单数量', table_style)
x += 2
for key in POofSupplier:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(POofSupplier[key]), text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年%d' % currentMonth + '月供应商份额统计', header_style)
newsheet.write(x+1, y+0, '供应商名称', table_style)
newsheet.write(x+1, y+1, '订单数量', table_style)
x += 2
for key in SePOofSupplier:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(SePOofSupplier[key]), text_style)
    x += 1


x += 3
y = 0
newsheet.write(x+0, y+0, '累计供应商各省分覆盖率情况', header_style)
newsheet.write(x+1, y+0, '供应商名称', table_style)
newsheet.write(x+1, y+1, '覆盖省分数', table_style)
x += 2
for key in TotalProvinceAndSupplier:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(TotalProvinceAndSupplier[key]), text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年供应商各省分覆盖率情况', header_style)
newsheet.write(x+1, y+0, '供应商名称', table_style)
newsheet.write(x+1, y+1, '覆盖省分数', table_style)
x += 2
for key in ProvinceAndSupplier:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(ProvinceAndSupplier[key]), text_style)
    x += 1


x += 3
y = 0
newsheet.write(x+0, y+0, '2017年%d' % currentMonth + '月供应商各省分覆盖率情况', header_style)
newsheet.write(x+1, y+0, '供应商名称', table_style)
newsheet.write(x+1, y+1, '覆盖省分数', table_style)
x += 2
for key in SeProvinceAndSupplier:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(SeProvinceAndSupplier[key]), text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '累计省分采购覆盖供应商情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '覆盖供应商数', table_style)
x += 2
for key in TotalSupplierOfPro:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(TotalSupplierOfPro[key]), text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年省分采购覆盖供应商情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '覆盖供应商数', table_style)
x += 2
for key in SupplierOfPro:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(SupplierOfPro[key]), text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年%d' % currentMonth + '月省分采购覆盖供应商情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '覆盖供应商数', table_style)
x += 2
for key in SeSupplierOfPro:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, len(SeSupplierOfPro[key]), text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年到货时间情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '最大到货时间', table_style)
newsheet.write(x+1, y+2, '最小到货时间', table_style)
newsheet.write(x+1, y+3, '平均到货时间', table_style)
x += 2
for key in ProvinceAndArrivalDay:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, max(ProvinceAndArrivalDay[key]), text_style)
    newsheet.write(x, y + 2, min(ProvinceAndArrivalDay[key]), text_style)
    newsheet.write(x, y + 3, round(sum(ProvinceAndArrivalDay[key])/len(ProvinceAndArrivalDay[key])), text_style)
    x += 1

x += 3
y = 0
newsheet.write(x+0, y+0, '2017年%d' % currentMonth + '月到货时间情况', header_style)
newsheet.write(x+1, y+0, '省分', table_style)
newsheet.write(x+1, y+1, '最大到货时间', table_style)
newsheet.write(x+1, y+2, '最小到货时间', table_style)
newsheet.write(x+1, y+3, '平均到货时间', table_style)
x += 2
for key in SeProvinceAndArrivalDay:
    newsheet.write(x, y, key, text_style)
    newsheet.write(x, y+1, max(SeProvinceAndArrivalDay[key]), text_style)
    newsheet.write(x, y + 2, min(SeProvinceAndArrivalDay[key]), text_style)
    newsheet.write(x, y + 3, round(sum(SeProvinceAndArrivalDay[key])/len(SeProvinceAndArrivalDay[key])), text_style)
    x += 1

wb.save(file_path)

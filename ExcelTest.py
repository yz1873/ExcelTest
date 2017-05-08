import xlrd
import sys

# 当月（2017年几月）
currentMonth = 4

data = xlrd.open_workbook(r"C:\Users\Zhang Yu\Desktop\微机站订单（基于分类+协议）.xlsx")
# 路径前加r，原因：文件名中的 \U 开始的字符被编译器认为是八进制
booksheet = data.sheets()[0] # 打开第一张表

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

            # 2017年统计  4月的没算???
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

# 由（供应商-省分）字典转化为（省分-供应商）字典
dictTransform(ProvinceAndSupplier,SupplierOfPro)
dictTransform(SeProvinceAndSupplier,SeSupplierOfPro)
dictTransform(TotalProvinceAndSupplier,TotalSupplierOfPro)


print("-----------------------------------------------------------------------------")
print("累计交易情况：   2015年订单数为：",len(POOfY[2015]),"     2016年订单数为：",len(POOfY[2016]),"     2017年订单数为：",len(POOfY[2017]))
print("-----------------------------------------------------------------------------")
print("2015年省分下单情况")
for key in POOf2015:
    print("第%d"%key,"个月的订单量为：%-5d"%len(POOf2015[key]),"  交易额为：",ListOfTotalPrice2015[key-1])
print("-----------------------------------------------------------------------------")
print("2016年省分下单情况")
for key in POOf2016:
    print("第%d"%key,"个月的订单量为：%-5d"%len(POOf2016[key]),"  交易额为：",ListOfTotalPrice2016[key-1])
print("-----------------------------------------------------------------------------")
print("2017年省分下单情况")
for key in POOf2017:
    print("第%d"%key,"个月的订单量为：%-5d"%len(POOf2017[key]),"  交易额为：",ListOfTotalPrice2017[key-1])
print("-----------------------------------------------------------------------------")
print("省分累计下单情况")
for key in TotalProAndPO:
    print(key, "--订单量为:%-5d"%len(TotalProAndPO[key]),"  交易额为：",TotalProAndTotalPrice[key])
print("-----------------------------------------------------------------------------")
print("2017年省分下单情况")
for key in ProAndPO:
    print(key, "--订单量为:%-5d"%len(ProAndPO[key]),"  交易额为：",ProAndTotalPrice[key])
print("-----------------------------------------------------------------------------")
print("2017年%d月省分下单情况"%currentMonth)
for key in SeProAndPO:
    print(key, "--订单量为:%-5d"%len(SeProAndPO[key]),"  交易额为：",SeProAndTotalPrice[key])
print("-----------------------------------------------------------------------------")

print("省分累计典配模式下单情况")
for key in CaTotalProAndPO:
    print(key, "典配订单量为:", len(CaTotalProAndPO[key]))
print("-----------------------------------------------------------------------------")
print("2017年省分累计典配模式下单情况")
for key in CaProAndPO:
    print(key, "典配订单量为:", len(CaProAndPO[key]))
print("-----------------------------------------------------------------------------")
print("2017年%d月省分累计典配模式下单情况"%currentMonth)
for key in CaSeProAndPO:
    print(key, "典配订单量为:", len(CaSeProAndPO[key]))
print("-----------------------------------------------------------------------------")

print("累计供应商订单数")
for key in TotalPOofSupplier:
    print("%-25s"%key, "订单量为:", len(TotalPOofSupplier[key]))
print("-----------------------------------------------------------------------------")
print("2017年供应商订单数")
for key in POofSupplier:
    print("%-25s"%key, "订单量为:", len(POofSupplier[key]))
print("-----------------------------------------------------------------------------")
print("2017年%d月供应商订单数"%currentMonth)
for key in SePOofSupplier:
    print("%-25s"%key, "订单量为:", len(SePOofSupplier[key]))
print("-----------------------------------------------------------------------------")


print("累计供应商覆盖省分")
for key in TotalProvinceAndSupplier:
    print("%-25s"%key, "覆盖省分数为:", len(TotalProvinceAndSupplier[key]), "覆盖省分为:", TotalProvinceAndSupplier[key])
print("-----------------------------------------------------------------------------")
print("2017年供应商覆盖省分")
for key in ProvinceAndSupplier:
    print("%-25s"%key, "覆盖省分数为:", len(ProvinceAndSupplier[key]), "覆盖省分为:", ProvinceAndSupplier[key])
print("-----------------------------------------------------------------------------")
print("2017年%d月供应商覆盖省分"%currentMonth)
for key in SeProvinceAndSupplier:
    print("%-25s"%key, "覆盖省分数为:", len(SeProvinceAndSupplier[key]), "覆盖省分为:", SeProvinceAndSupplier[key])
print("-----------------------------------------------------------------------------")

print("累计省分采购覆盖供应商")
for key in TotalSupplierOfPro:
    print("%-5s"%key,"覆盖供应商数为:", len(TotalSupplierOfPro[key]), "覆盖供应商为:", TotalSupplierOfPro[key])
print("-----------------------------------------------------------------------------")
print("2017年省分采购覆盖供应商")
for key in SupplierOfPro:
    print("%-5s"%key,"覆盖供应商数为:", len(SupplierOfPro[key]), "覆盖供应商为:", SupplierOfPro[key])
print("-----------------------------------------------------------------------------")
print("2017年%d月省分采购覆盖供应商"%currentMonth)
for key in SeSupplierOfPro:
    print("%-5s"%key,"覆盖供应商数为:", len(SeSupplierOfPro[key]), "覆盖供应商为:", SeSupplierOfPro[key])
print("-----------------------------------------------------------------------------")

print("2017年到货时间情况")
for key in ProvinceAndArrivalDay:
    print("%-5s"%key, "最大到货时间为:%-5d"%max(ProvinceAndArrivalDay[key]),"最小到货时间为:%-5d"%min(ProvinceAndArrivalDay[key]),
          "平均到货时间为:%-5d"%(sum(ProvinceAndArrivalDay[key])/len(ProvinceAndArrivalDay[key])))
print("-----------------------------------------------------------------------------")

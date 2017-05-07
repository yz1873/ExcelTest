import xlrd
import sys

#当月（2017年几月）
currentMonth = 3

data = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\1.xlsx")
#路径前加r，原因：文件名中的 \U 开始的字符被编译器认为是八进制
booksheet = data.sheets()[0] # 打开第一张表

#找到订单编号列
colPO = -1
#找到下单时间列
colPlaceTime = -1
#找到订单状态列
colOrderStatus = -1
#找到单个净价合计列
colTotalPrice = -1
#找到省份名称列
colProvince = -1
#找到分类列
colCategory = -1
#找到供应商列
colSupplier = -1

#python3.0后 xrange和range一样
for col in range(booksheet.ncols):
    cell = booksheet.cell_value(0, col)
    if cell == "订单编号":
        colPO = col
    if cell == "下单时间":
        colPlaceTime = col
    if cell == "订单状态":
        colOrderStatus = col
    if cell == "单个净价合计":
        colTotalPrice = col
    if cell == "省份名称":
        colProvince = col
    if cell == "分类":
        colCategory = col
    if cell == "供应商":
        colSupplier = col
if colPO == -1:
    print("没有找到订单编号列！")
    sys.exit(0)
if colPlaceTime == -1:
    print("没有找到下单时间列！")
    sys.exit(0)
if colOrderStatus == -1:
    print("没有找到订单状态列！")
    sys.exit(0)
if colTotalPrice == -1:
    print("没有找到单个净价合计列！")
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

#2016与2017年订单集合
POOfY = {2016:set(),2017:set()}
# index从0-7分别代表2016年9月到2017年4月  各月净价总和
ListOfTotalPrice = [0,0,0,0,0,0,0,0]
#各个月对应订单集合
POOfYM = {201609:set(),201610:set(),201611:set(),201612:set(),201701:set(),201702:set(),201703:set(),201704:set()}

#以下字典key均为省分名称
#2017年省分与订单集合的字典
ProAndPO = {}
#2017年省分与交易额的字典
ProAndTotalPrice = {}
#特殊月份省分与订单集合的字典
SeProAndPO = {}
#特殊月份省分与交易额的字典
SeProAndTotalPrice = {}
#2016年与2017年总省分与订单集合的字典
TotalProAndPO = {}
#2016年与2017年总省分与交易额的字典
TotalProAndTotalPrice = {}
#典配—2017年省分与订单集合的字典
CaProAndPO = {}
#典配—特殊月份省分与订单集合的字典
CaSeProAndPO = {}
#典配—2016年与2017年总省分与订单集合的字典
CaTotalProAndPO = {}
#本省2017年的供应商
SupplierOfPro = {}
#本省2017年特殊月的供应商
SeSupplierOfPro = {}

#以下字典key均为供应商名称
#2017年供应商与订单集合字典
POofSupplier = {}
#2017年特殊月供应商与订单集合字典
SePOofSupplier = {}
#2017年供应商与省分集合字典
ProvinceAndSupplier = {}
#2017年特殊月供应商与省分集合字典
SeProvinceAndSupplier = {}



for row in range(booksheet.nrows):
    if row == 0:  # 跳过第一行
        continue
    if (booksheet.cell_value(row, colOrderStatus)=="确认订单" or booksheet.cell_value(row, colOrderStatus)=="完成订单"):
        date_value = xlrd.xldate_as_tuple(booksheet.cell_value(row, colPlaceTime), data.datemode)
        if date_value[0] == 2016:
            POOfY[2016].add(booksheet.cell_value(row, colPO))
            if date_value[1]==9:
                ListOfTotalPrice[0] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201609].add(booksheet.cell_value(row, colPO))
            if date_value[1]==10:
                ListOfTotalPrice[1] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201610].add(booksheet.cell_value(row, colPO))
            if date_value[1]==11:
                ListOfTotalPrice[2] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201611].add(booksheet.cell_value(row, colPO))
            if date_value[1]==12:
                ListOfTotalPrice[3] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201612].add(booksheet.cell_value(row, colPO))
        if date_value[0] == 2017:
            POOfY[2017].add(booksheet.cell_value(row, colPO))
            if date_value[1]==1:
                ListOfTotalPrice[4] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201701].add(booksheet.cell_value(row, colPO))
            if date_value[1]==2:
                ListOfTotalPrice[5] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201702].add(booksheet.cell_value(row, colPO))
            if date_value[1]==3:
                ListOfTotalPrice[6] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201703].add(booksheet.cell_value(row, colPO))
            if date_value[1]==4:
                ListOfTotalPrice[7] += booksheet.cell_value(row, colTotalPrice)
                POOfYM[201704].add(booksheet.cell_value(row, colPO))
            #2017年统计  4月的没算???
            # if (date_value[1] == 1 or date_value[1]==2 or date_value[1]==3):
            #省分订单与总价统计
            if (booksheet.cell_value(row, colProvince)[0:2] not in ProAndPO):
                temp = set()
                temp.add(booksheet.cell_value(row, colPO))
                ProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
                ProAndTotalPrice[booksheet.cell_value(row, colProvince)[0:2]] = booksheet.cell_value(row, colTotalPrice)
            else:
                temp = ProAndPO[booksheet.cell_value(row, colProvince)[0:2]]
                temp.add(booksheet.cell_value(row, colPO))
                ProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
                ProAndTotalPrice[booksheet.cell_value(row, colProvince)[0:2]] += booksheet.cell_value(row, colTotalPrice)

                #统计典配
            if ("组合商品" in booksheet.cell_value(row, colCategory)):
                if (booksheet.cell_value(row, colProvince)[0:2] not in CaProAndPO):
                    temp = set()
                    temp.add(booksheet.cell_value(row, colPO))
                    CaProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
                else:
                    temp = CaProAndPO[booksheet.cell_value(row, colProvince)[0:2]]
                    temp.add(booksheet.cell_value(row, colPO))
                    CaProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp

            # 2017年供应商维度统计
            if (booksheet.cell_value(row, colSupplier) not in POofSupplier):
                temp = set()
                temp.add(booksheet.cell_value(row, colPO))
                POofSupplier[booksheet.cell_value(row, colSupplier)] = temp
                stemp = set()
                stemp.add(booksheet.cell_value(row, colProvince)[0:2])
                ProvinceAndSupplier[booksheet.cell_value(row, colSupplier)] = stemp
            else:
                temp = POofSupplier[booksheet.cell_value(row, colSupplier)]
                temp.add(booksheet.cell_value(row, colPO))
                POofSupplier[booksheet.cell_value(row, colSupplier)] = temp
                stemp = ProvinceAndSupplier[booksheet.cell_value(row, colSupplier)]
                stemp.add(booksheet.cell_value(row, colProvince)[0:2])
                ProvinceAndSupplier[booksheet.cell_value(row, colSupplier)] = stemp

            #特殊月为当月 特殊月统计
            if (date_value[1]==currentMonth):
                # 省分订单与总价统计
                if (booksheet.cell_value(row, colProvince)[0:2] not in SeProAndPO):
                    temp = set()
                    temp.add(booksheet.cell_value(row, colPO))
                    SeProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
                    SeProAndTotalPrice[booksheet.cell_value(row, colProvince)[0:2]] = booksheet.cell_value(row, colTotalPrice)
                else:
                    temp = SeProAndPO[booksheet.cell_value(row, colProvince)[0:2]]
                    temp.add(booksheet.cell_value(row, colPO))
                    SeProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
                    SeProAndTotalPrice[booksheet.cell_value(row, colProvince)[0:2]] += booksheet.cell_value(row, colTotalPrice)

                #此处与致宏数据有所冲突，原因为她在订单删除重复项时，删掉了带有“组合商品”字样的订单，而保留了“配套服务”
                #其实该订单依然为典配。
                #典配统计
                if ("组合商品" in booksheet.cell_value(row, colCategory)):
                    if (booksheet.cell_value(row, colProvince)[0:2] not in CaSeProAndPO):
                        temp = set()
                        temp.add(booksheet.cell_value(row, colPO))
                        CaSeProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
                    else:
                        temp = CaSeProAndPO[booksheet.cell_value(row, colProvince)[0:2]]
                        temp.add(booksheet.cell_value(row, colPO))
                        CaSeProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp

                # 供应商维度统计
                if (booksheet.cell_value(row, colSupplier) not in SePOofSupplier):
                    temp = set()
                    temp.add(booksheet.cell_value(row, colPO))
                    SePOofSupplier[booksheet.cell_value(row, colSupplier)] = temp
                    stemp = set()
                    stemp.add(booksheet.cell_value(row, colProvince)[0:2])
                    SeProvinceAndSupplier[booksheet.cell_value(row, colSupplier)] = stemp
                else:
                    temp = SePOofSupplier[booksheet.cell_value(row, colSupplier)]
                    temp.add(booksheet.cell_value(row, colPO))
                    SePOofSupplier[booksheet.cell_value(row, colSupplier)] = temp
                    stemp = SeProvinceAndSupplier[booksheet.cell_value(row, colSupplier)]
                    stemp.add(booksheet.cell_value(row, colProvince)[0:2])
                    SeProvinceAndSupplier[booksheet.cell_value(row, colSupplier)] = stemp

        #无论年份
        # 省分订单与总价统计
        if (booksheet.cell_value(row, colProvince)[0:2] not in TotalProAndPO):
            temp = set()
            temp.add(booksheet.cell_value(row, colPO))
            TotalProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
            TotalProAndTotalPrice[booksheet.cell_value(row, colProvince)[0:2]] = booksheet.cell_value(row, colTotalPrice)
        else:
            temp = TotalProAndPO[booksheet.cell_value(row, colProvince)[0:2]]
            temp.add(booksheet.cell_value(row, colPO))
            TotalProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
            TotalProAndTotalPrice[booksheet.cell_value(row, colProvince)[0:2]] += booksheet.cell_value(row, colTotalPrice)

        #典配统计
        if ("组合商品" in booksheet.cell_value(row, colCategory)):
            if (booksheet.cell_value(row, colProvince)[0:2] not in CaTotalProAndPO):
                temp = set()
                temp.add(booksheet.cell_value(row, colPO))
                CaTotalProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp
            else:
                temp = CaTotalProAndPO[booksheet.cell_value(row, colProvince)[0:2]]
                temp.add(booksheet.cell_value(row, colPO))
                CaTotalProAndPO[booksheet.cell_value(row, colProvince)[0:2]] = temp

#由（供应商-省分）字典转化为（省分-供应商）字典
for key in ProvinceAndSupplier:
    for p in ProvinceAndSupplier[key]:
        if p not in SupplierOfPro:
            temp = set()
            temp.add(key)
            SupplierOfPro[p] = temp
        else:
            temp = SupplierOfPro[p]
            temp.add(key)
            SupplierOfPro[p] = temp
for key in SeProvinceAndSupplier:
    for p in SeProvinceAndSupplier[key]:
        if p not in SeSupplierOfPro:
            temp = set()
            temp.add(key)
            SeSupplierOfPro[p] = temp
        else:
            temp = SeSupplierOfPro[p]
            temp.add(key)
            SeSupplierOfPro[p] = temp

print("-----------------------------------------------------------------------------")
print("累计交易情况：2016年订单数为：",len(POOfY[2016])," 2017年订单数为：",len(POOfY[2017]))
print("-----------------------------------------------------------------------------")
print("2016年9月到2017年4月订单量为：",[len(POOfYM[201609]),len(POOfYM[201610]),len(POOfYM[201611]),
                              len(POOfYM[201612]),len(POOfYM[201701]),len(POOfYM[201702]),len(POOfYM[201703]),len(POOfYM[201704])])
print("-----------------------------------------------------------------------------")
print("2016年9月到2017年4月交易额为：",ListOfTotalPrice)
print("-----------------------------------------------------------------------------")
print("2017年省分下单情况")
for key in ProAndPO:
    print(key, "--订单量为:%-5d"%len(ProAndPO[key]),"  交易额为：",ProAndTotalPrice[key])
print("-----------------------------------------------------------------------------")
print("2017年当月省分下单情况")
for key in SeProAndPO:
    print(key, "--订单量为:%-5d"%len(SeProAndPO[key]),"  交易额为：",SeProAndTotalPrice[key])
print("-----------------------------------------------------------------------------")
print("省分累计下单情况")
for key in TotalProAndPO:
    print(key, "--订单量为:%-5d"%len(TotalProAndPO[key]),"  交易额为：",TotalProAndTotalPrice[key])
print("-----------------------------------------------------------------------------")

print("省分累计典配模式下单情况")
for key in CaTotalProAndPO:
    print(key, "典配订单量为:", len(CaTotalProAndPO[key]))
print("-----------------------------------------------------------------------------")
print("2017年省分累计典配模式下单情况")
for key in CaProAndPO:
    print(key, "典配订单量为:", len(CaProAndPO[key]))
print("-----------------------------------------------------------------------------")
print("2017年当月省分累计典配模式下单情况")
for key in CaSeProAndPO:
    print(key, "典配订单量为:", len(CaSeProAndPO[key]))
print("-----------------------------------------------------------------------------")

print("2017年供应商订单数")
for key in POofSupplier:
    print("%-25s"%key, "订单量为:", len(POofSupplier[key]))
print("-----------------------------------------------------------------------------")
print("2017年当月供应商订单数")
for key in SePOofSupplier:
    print("%-25s"%key, "订单量为:", len(SePOofSupplier[key]))
print("-----------------------------------------------------------------------------")


print("2017年供应商覆盖省分")
for key in ProvinceAndSupplier:
    print("%-25s"%key, "覆盖省分为:", ProvinceAndSupplier[key])
print("-----------------------------------------------------------------------------")
print("2017年当月供应商覆盖省分")
for key in SeProvinceAndSupplier:
    print("%-25s"%key, "覆盖省分为:", SeProvinceAndSupplier[key])
print("-----------------------------------------------------------------------------")

print("2017年省分采购覆盖供应商")
for key in SupplierOfPro:
    print("%-5s"%key, "覆盖供应商为:", SupplierOfPro[key])
print("-----------------------------------------------------------------------------")
print("2017年当月省分采购覆盖供应商")
for key in SeSupplierOfPro:
    print("%-5s"%key, "覆盖供应商为:", SeSupplierOfPro[key])
print("-----------------------------------------------------------------------------")


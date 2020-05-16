import datetime

from openpyxl import load_workbook, Workbook


#判断是否在活动区间
def checkTimeRange(tmpPddPayTime, tmpBeginTime, tmpEndTime, range):
     tmpEndTimePlus = tmpEndTime+datetime.timedelta(seconds=range)
     result = tmpPddPayTime >= tmpBeginTime and tmpPddPayTime <= tmpEndTimePlus
     return result



#生成excel
#dict[(sku,活动编号，成本价格)]{(商品名称,数量)}
def createExcel(outFilepath,dictInfo):
    newWb = Workbook()

    newSheet = newWb.active

    # 填写字段抬头
    beginRow = 1
    newSheet.cell(row=beginRow, column=1, value="sku")
    newSheet.cell(row=beginRow, column=2, value="活动编号")
    newSheet.cell(row=beginRow, column=3, value="价格")
    newSheet.cell(row=beginRow, column=4, value="商品名称")
    newSheet.cell(row=beginRow, column=5, value="数量")
    beginRow = beginRow + 1

    for key in dictInfo:
        sku = key[0]
        actyCode = key[1]
        basicPrice = key[2]
        proName = dictInfo[key][0]
        num = dictInfo[key][1]

        newSheet.cell(row=beginRow, column=1, value=sku)
        newSheet.cell(row=beginRow, column=2, value=actyCode)
        newSheet.cell(row=beginRow, column=3, value=basicPrice)
        newSheet.cell(row=beginRow, column=4, value=proName)
        newSheet.cell(row=beginRow, column=5, value=num)

        beginRow = beginRow + 1

    newWb.save(outFilepath)


#读取截单数据，输出list[原始单号]
def getJieDanExcel(excelPath,originalOrderNumPos,filterInfo,beginRow):
    workBook = load_workbook(excelPath)
    sheet = workBook.active
    rows = sheet.max_row
    listToReturn = list()

    #遍历每行读取数据
    i = beginRow
    while i <= rows:
        listToReturn.append(str(sheet.cell(row=i, column=originalOrderNumPos).value).strip().strip(filterInfo))
        i = i + 1
    return listToReturn


#读取拼多多表，截单list[原始单号]，根据原始单号过滤，输出list[tuple（订单号，pddIdPos,sku，支付时间，商品总价,商品名称）]
def getPddExcelByFilter(excelPath,beginRow,originalOrderNumList,pddOrderNumPos,pddIdPos,pddSkuPos,pddPayTimePos,productPricePos,proNamePos,filterInfo):
    workBook = load_workbook(excelPath)
    sheet = workBook.active
    rows = sheet.max_row
    listToReturn = list()

    #遍历每一行
    # 遍历每行读取数据
    i = beginRow
    while i <= rows and len(originalOrderNumList)>0:
        print("拼多多数据处理中： "+str(i)+"/"+str(rows))
        #遍历截单的订单号
        j = 0
        while j < len(originalOrderNumList):
            tmpOrderNum = str(sheet.cell(row=i, column=pddOrderNumPos).value.strip().strip(filterInfo))
            # 判断订单号是否在list中
            if str(originalOrderNumList[j]) == tmpOrderNum:
                print("截单数据整理中:"+str(j)+"/"+str(len(originalOrderNumList)))
                tmpPddId = str(sheet.cell(row=i, column=pddIdPos).value).strip().strip(filterInfo)
                tmpSku = str(sheet.cell(row=i, column=pddSkuPos).value).strip().strip(filterInfo)
                tmpPayTime = sheet.cell(row=i, column=pddPayTimePos).value
                tmpProductPrice = float(str(sheet.cell(row=i, column=productPricePos).value).strip().strip(filterInfo))
                tmpProName = str(sheet.cell(row=i, column=proNamePos).value).strip().strip(filterInfo)
                listToReturn.append((tmpOrderNum,tmpPddId,tmpSku,tmpPayTime,tmpProductPrice,tmpProName))

                #从list中删除以处理的数据，缩短丽水
                originalOrderNumList.pop(j)
                #跳出循环
                break
            j = j + 1
        i = i + 1

    print("未处理数据数量：")
    print(originalOrderNumList)
    return listToReturn


#读取活动备案表，输出list[tuple(活动编号,开始时间，结束时间,tuple(pddId,sku,活动价，采购价)]
def getSecctionExcel(excelPath,sheetName,beginRow,secctionCodePos,secctionBeginTimePos,secctionEndTimePos,secctionPddIdPos,secctionSkuPos,secctionPricePos,secctionPurchasePricePos,filterInfo):
    workBook = load_workbook(excelPath)
    sheet = workBook[sheetName]
    rows = sheet.max_row
    columns = sheet.max_column
    listToReturn = list()
    print(rows)


    # 遍历每行读取数据
    i = beginRow
    while i <= rows:
        j = 1
        print("活动备案数据处理中： " + str(i) + "/" + str(rows))
        # while j <= columns:
        #如果是第一次遇到当前活动
        tmpSecctionCode = str(sheet.cell(row=i, column=secctionCodePos).value).strip().strip(filterInfo)
        tmpSecctionBeginTime = sheet.cell(row=i, column=secctionBeginTimePos).value

        if tmpSecctionBeginTime is None:
            tmpSecctionBeginTime =  datetime.datetime(2099, 9, 1)

        tmpSecctionEndTime = sheet.cell(row=i, column=secctionEndTimePos).value
        if tmpSecctionEndTime is None:
            tmpSecctionEndTime =  datetime.datetime(2099, 9, 1)

        listToReturn.append((
        tmpSecctionCode,tmpSecctionBeginTime,tmpSecctionEndTime,
        str(sheet.cell(row=i, column=secctionPddIdPos).value).strip().strip(filterInfo),
        str(sheet.cell(row=i, column=secctionSkuPos).value).strip().strip(filterInfo),
        sheet.cell(row=i, column=secctionPricePos).value,
        sheet.cell(row=i, column=secctionPurchasePricePos).value
        ))


        i = i + 1

    return listToReturn

#param：活动备案信息， pdd 订单信息，返回list[(订单编号，活动编号，sku，成本价格，商品名称)]
def runFinalWork(actSecctionList,pdd,range):
    pddSize = len(pdd) #需要处理订单量
    toReturn = list()


    count = 0
    while count < pddSize:
        toNext = 0 #如果已经找到对应商品，就标志

        print("匹配拼多多信息："+str(count)+"/"+str(pddSize))
        ordrNum = pdd[count][0] #订单编号
        tmpPddId = pdd[count][1] #商品id编码
        tmpPddSku  = pdd[count][2] #商品sku编码
        tmpPddPayTime = pdd[count][3] #支付时间
        tmpPddProductAllPrice = pdd[count][4] #商品总价
        tmpPddProName = pdd[count][5] #商品名称


        #遍历所有活动内容
        #list[tuple(活动编号,开始时间，结束时间,pddId,sku,活动价，采购价)]
        actSheetPos = 1
        for item in actSecctionList:
             actSheetPos = actSheetPos + 1
             if toNext == 1:
                 break

             #活动编码
             tmpActCode = item[0]
             #开始时间
             tmpBeginTime = item[1]
             #结束时间
             tmpEndTime = item[2]
             #pdd 的商品 id
             tmpPddId_actSheet = item[3]
             #商品的sku编码
             tmpSku_actSheet = item[4]
             #活动价
             tmpActPrice = item[5]
             #成本价
             tmpPurchasePrice = item[6]



            #判断是否在活动区间
             if checkTimeRange(tmpPddPayTime,tmpBeginTime,tmpEndTime,range):
                #如果是iphone
                if "iPad" in tmpPddProName:
                    # 是否是: 同样的sku &&
                    if tmpSku_actSheet == tmpPddSku:
                        # 活动价等于商品总价
                        if tmpActPrice == tmpPddProductAllPrice:
                            # sku
                            skuToReturn = tmpSku_actSheet
                            # 成本价格
                            basePriceToReturn = tmpPurchasePrice
                            # 商品名称
                            productNameToReturn = tmpPddProName
                            toReturn.append((ordrNum,tmpActCode, skuToReturn, basePriceToReturn, productNameToReturn))
                            #跳出活动表循环，准备匹配下一个订单
                            toNext = 1
                            break

                elif "iPhone" in tmpPddProName:
                    # 是否是: 同样的sku && 活动价等于商品总价 or 商品id一致(此处的or到时候要改成and)
                    if tmpPddId_actSheet == tmpPddId and (tmpSku_actSheet == tmpPddSku and tmpActPrice == tmpPddProductAllPrice):
                        # sku
                        skuToReturn = tmpSku_actSheet
                        # 成本价格
                        basePriceToReturn = tmpPurchasePrice
                        # 商品名称
                        productNameToReturn = tmpPddProName
                        toReturn.append((ordrNum,tmpActCode, skuToReturn, basePriceToReturn, productNameToReturn))
                        toNext = 1
                        break
                #如果是其他
                else:
                    print("其他商品："+tmpPddProName)

        count = count + 1
    return toReturn


#商品分类汇总计算数量。 param: list[(订单编号，活动编号，sku，成本价格，商品名称)]; return: dict[(sku,活动编号，成本价格)]{(商品名称,数量)}
def clssify(purchaseList):
    dictToReturn = dict()
    for item in purchaseList:
        tmpActCode = item[1]
        tmpSku = item[2]
        tmpBasicPrice = item[3]
        tmpProName = item[4]

        if (tmpSku,tmpActCode,tmpBasicPrice) not in dictToReturn:
            dictToReturn[(tmpSku,tmpActCode,tmpBasicPrice)] = (tmpProName,1)
        else:
            oldInfo = dictToReturn[(tmpSku,tmpActCode,tmpBasicPrice)]
            dictToReturn[(tmpSku, tmpActCode, tmpBasicPrice)] = (oldInfo[0], oldInfo[1]+1)

    return dictToReturn

#根据订单号计算不符合标准订单,param  : list[tuple（订单号，pddIdPos,sku，支付时间，商品总价,商品名称）
def checkList(pddListBefore,pddListAfter):
    listToReturn = list()
    for oldItem in pddListBefore:
        #判断老的信息是否在新的列表里出现
        isTreated = 0
        pddListAfterSize = len(pddListAfter)
        index = 0
        while index < pddListAfterSize:
            newItem = pddListAfter[index]
            if newItem[0] == oldItem[0]:
                isTreated = 1
                break
            index = index + 1
        if isTreated == 0:
            listToReturn.append(oldItem)

    return listToReturn


jieDanExcelPath = "C:/Users/admin/Desktop/截单/jiedan.xlsx"
jieDanOriginalOrderNumPos = 3
beginRow = 2
filterInfo = "\t"

jieDanList = getJieDanExcel(jieDanExcelPath,jieDanOriginalOrderNumPos,filterInfo,beginRow)
jieDanListSize = len(jieDanList)

pddExcelPath = "C:/Users/admin/Desktop/截单/pddTest.xlsx"
originalOrderNumList = jieDanList
pddOrderNumPos = 2
pddSkuPos = 33
pddPayTimePos = 23
productPricePos = 4
pddIdPos = 29
proNamePos = 1
pddFilterInfo = "\t"
beginRow = 2
pddExcel = getPddExcelByFilter(pddExcelPath,beginRow,originalOrderNumList,pddOrderNumPos,pddIdPos,pddSkuPos,pddPayTimePos,productPricePos,proNamePos,pddFilterInfo)


pddFinanceExcelPath = "C:/Users/admin/Desktop/截单/活动备案.xlsx"
sheetName = "活动备案"
secctionCodePos = 3
secctionBeginTimePos = 4
secctionEndTimePos = 5
secctionSkuPos = 6
secctionPricePos = 11
secctionPurchasePricePos = 9
secctionPddIdPos = 2
filterInfo = "\t"
beginRow = 2
financeDict = getSecctionExcel(pddFinanceExcelPath,sheetName,beginRow,secctionCodePos,secctionBeginTimePos,secctionEndTimePos,secctionPddIdPos,secctionSkuPos,secctionPricePos,secctionPurchasePricePos,filterInfo)
print(financeDict)


# tmpFinanceDict =
# tmpPddExcel =
#活动浮动区间
range = 18000
purchaseExcel=runFinalWork(financeDict,pddExcel,range)

#检查未符合标准的订单
faultList = checkList(pddExcel,purchaseExcel)

#分类汇总数据
finalInfo = clssify(purchaseExcel)

#标记pdd Excel 中订单对应的活动编码
def markActCode(purchaseExcel, pddExcelPath,ordNumPos,toWritePos,outputPath):
    workBook = load_workbook(pddExcelPath)
    sheet = workBook.active
    rows = sheet.max_row
    beginRow = 2
    i = beginRow
    while i <= rows:
        ordNum = str(sheet.cell(row=i,column=ordNumPos).value.strip().strip(filterInfo))
        #list[(订单编号，活动编号，sku，成本价格，商品名称)]
        for item in purchaseExcel:
            if item[0] == ordNum:
                sheet.cell(row=i, column=toWritePos,value=item[1])
                break
        i = i + 1
    workBook.save(outputPath)

toWritePos = 56
outputPath = "C:/Users/admin/Desktop/截单/pddTestWithActCode.xlsx"
markActCode(purchaseExcel,pddExcelPath,pddOrderNumPos,toWritePos,outputPath)

print("________备案表获取结果 financeDict________________")
print(len(financeDict))
print(financeDict)
print()

print("_________截单表内不符合拼多多订单表内信息 jieDanList______________")
print(jieDanList)
print()

print("_________拼多多过滤截单后数据 pddExcel_____________")
print(len(pddExcel))
print(pddExcel)
print()

print("____________对比备案表结果 finalExcel________")
print(len(purchaseExcel))
print(purchaseExcel)
print()

print("______________未符合标准商品:________")
print(len(faultList))
print(faultList)
print()

print("______________分类汇总最终结果 finalInfo________")
print(len(finalInfo))
print(finalInfo)
print()



outFilepath = "C:/Users/admin/Desktop/test/outPut.xlsx"
createExcel(outFilepath,finalInfo)
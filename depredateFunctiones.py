def runFinalWork_Depredate(actSecctionDict,pdd):
    #活动编号
    actToReturn = None
    #sku
    skuToReturn = None
    #成本价格
    basePriceToReturn = None
    #商品名称
    productNameToReturn = None
    toReturn = list()

    for itemKey in pdd:
        tmpPddSku  = pdd[itemKey][0] #商品sku编码
        tmpPddPayTime = pdd[itemKey][1].timestamp() #支付时间
        tmpPddProductAllPrice = pdd[itemKey][2] #商品总价

        #遍历所有活动内容
        for actName in actSecctionDict:
             tmpInfo = actSecctionDict[actName]
             tmpBeginTime = tmpInfo[0].timestamp()
             tmpEndTime = tmpInfo[1].timestamp()
             tmpSkuInfoList = tmpInfo[2]

             #是否在活动时间之内
             if tmpPddPayTime >= tmpBeginTime and tmpPddPayTime <= tmpEndTime:
                # 活动编号
                actToReturn = actName
                #遍历参加活动的所有商品
                for skuInfo in tmpSkuInfoList:
                    #是否是同样的sku和活动价是否等于商品总价
                    if skuInfo[0] == tmpPddSku and  skuInfo[1] == tmpPddProductAllPrice:
                        # sku
                        skuToReturn = skuInfo
                        # 成本价格
                        basePriceToReturn = skuInfo[2]
                        # 商品名称
                        productNameToReturn = None

                        toReturn.append((actToReturn,skuToReturn,basePriceToReturn,productNameToReturn))

    return toReturn

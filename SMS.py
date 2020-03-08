import pandas as pd
from random import randint
import xlsxwriter
#buys 70 papers each day at 0.33
noOfPaper = 70
costOfPaper = 0.33
#sells paper at 0.5 and earns 0.05 for scrap
sellPrice = 0.5
scrapPrice = 0.05
colNames = ['Day','RDA by Day','Type of Day','RDA for demand','Demand','Revenue from Sales','Loss profit from excess','Scraps','Daily profit']
finalData = []
finalData.append(colNames)
summationRevenueSales=0
summationLossProfit = 0
summationScraps = 0
summationProfit = 0
for i in range(1,501):
    data=[]
    data.append(i)
    rdaDay = randint(0,100)
    data.append(rdaDay)
    typeDay=''
    if rdaDay >= 0 and rdaDay <= 40:
        typeDay = 'Good'
    elif rdaDay >= 41 and rdaDay <= 80:
        typeDay = 'Fair'
    else:
        typeDay = 'Poor'
    data.append(typeDay)
    rdaDemand = randint(0,100)
    data.append(rdaDemand)
    demand = randint(40,100)
    data.append(demand)
    revenueSales = -1
    if demand < noOfPaper :
        revenueSales = demand * sellPrice
    else:
        revenueSales = noOfPaper * sellPrice
    summationRevenueSales += revenueSales
    data.append(revenueSales)
    lossExcess = 0
    if demand > noOfPaper :
        lossExcess = (demand - noOfPaper) * 0.17
    summationLossProfit += lossExcess
    data.append(lossExcess)
    scraps = 0
    if demand < noOfPaper :
        scraps = (noOfPaper - demand) * scrapPrice
    summationScraps += scraps
    data.append(scraps)
    dailyProfit = revenueSales - (noOfPaper * costOfPaper ) - lossExcess + scraps
    summationProfit += dailyProfit
    data.append(dailyProfit)
    finalData.append(data)
df = pd.DataFrame(finalData)
writer = pd.ExcelWriter('SMS_Paper_Seller.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='DhanviPDesai',startrow=0, header=False,index=False)
workbook = writer.book
worksheet = writer.sheets['DhanviPDesai']
worksheet.write_number('F502',summationRevenueSales)
worksheet.write_number('G502',summationLossProfit)
worksheet.write_number('H502',summationScraps)
worksheet.write_number('I502',summationProfit)
workbook.close()
writer.save()

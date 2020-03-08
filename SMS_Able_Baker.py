import pandas as pd
from random import randint
import xlsxwriter
table1 = {'1':[0.25,0.25,0,25],'2':[0.40,0.65,26,65],'3':[0.2,0.85,66,85],'4':[0.15,1,86,100]}
able = {'2':[0.3,0.3,0,30],'3':[0.28,0.58,31,58],'4':[0.25,0.83,59,83],'5':[0.17,1,84,100]}
baker = {'3':[0.35,0.35,0,35],'4':[0.25,0.60,36,60],'5':[0.2,0.8,61,80],'6':[0.2,1,81,100]}
colnames = ['Cutomers','RAD for Arrival','Inter Arrival Time','Arrival Time','RAD for Service','Handled By','TSB','ST','TSE','Waiting Time','System Time']
def findHandledBy(at,abletse,bakertse):
    #return 0 for able and 1 for baker
    wt = 0
    if at >= abletse:
        return 0,wt
    elif at >= bakertse:
        return 1,wt
    else:
        diffab = abletse - at
        diffba = bakertse - at
        if diffba > diffab :
            wt = diffab
            return 0,wt
        else:
            wt = diffba
            return 1,wt
    
    
fullData = []
fullData.append(colnames)
prevat = -1
handledby = 0
abletsb = -1
bakertsb = -1
abletse = -1
bakertse = -1
SystemTime = 0
finalSystemTime = 0
waitingTime = 0
ableWorking = 0
bakerWorking = 0
totalWaitingTime = 0
for i in range(1,501):
    data = []
    data.append(i)
    at = 0
    randArrival = 0
    if i==1:
        randArrival = -1
    else:
        randArrival = randint(0,100)
    data.append(randArrival)
    iat = 0
    if randArrival == -1:
        iat = -1
        at=0
        prevat = at
        data.append(iat)
        data.append(at)
        handledBy = 0
    else:
        for key in table1.keys():
            currRow = table1[key]
            lowLimit = currRow[2]
            upLimit = currRow[3]
            if randArrival >= lowLimit and randArrival <= upLimit:
                iat = int(key)
        at = prevat + iat
        prevat = at
        data.append(iat)
        data.append(at)
        handledBy,waitingTime = findHandledBy(at,abletse,bakertse)
    service = randint(0,100)
    data.append(service)
    if handledBy == 0:
        data.append('Able')
        for key in able.keys():
            currRow = able[key]
            lowLimit = currRow[2]
            upLimit = currRow[3]
            if service >= lowLimit and service <= upLimit:
                st = int(key)
        data.append(at)
        data.append(st)
        abletse = at + st
        ableWorking += st
        data.append(abletse)
        SystemTime = st
    elif handledBy == 1:
        data.append('Baker')
        for key in baker.keys():
            currRow = baker[key]
            lowLimit = currRow[2]
            upLimit = currRow[3]
            if service >= lowLimit and service <= upLimit:
                st = int(key)
        data.append(at)
        data.append(st)
        bakertse = at + st
        bakerWorking += st
        data.append(bakertse)
        SystemTime = st
    totalWaitingTime += waitingTime
    data.append(waitingTime)
    data.append(SystemTime)
    finalSystemTime += SystemTime
    fullData.append(data)
df = pd.DataFrame(fullData)
writer = pd.ExcelWriter('SMS_Able_Baker.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='DhanviPDesai',startrow=0, header=False,index=False)
workbook = writer.book
ableIdle = finalSystemTime - ableWorking
bakerIdle = finalSystemTime - bakerWorking
worksheet = writer.sheets['DhanviPDesai']
worksheet.write('K502',finalSystemTime)
worksheet.write('J502',totalWaitingTime)
worksheet.write('F502','Able Idle Time = '+str(ableIdle))
worksheet.write('F503','Baker Idle Time = '+str(bakerIdle))
workbook.close()

import schedule
import time,datetime
from bll import Bllw
from ParseInvoiceWeb import ParseInvoiceWeb
from ParseInvoiceWeb import ParseFade
import json

ban=''
userid=''
password=''
companyname=''

with open('secret.txt','r',encoding='utf-8-sig') as f:
    payload=json.loads(f.read())
    ban=payload['ban']
    userid=payload['userid']
    password=payload['password']
    companyname=payload['companyname']
  
def GetNow():
    now=datetime.datetime.now()
    return now.strftime(r"%Y-%m-%d %H:%M:%S")

def CheckTemp():
    bll=Bllw()
    bll.CheckTemp()
    print ("check 機房溫度",GetNow())

def ClearMessageLog():
    bll=Bllw()
    bll.ClearMessageLog()
    print ("clear log",GetNow())

def RegisterDay():
    global ban,userid,password,companyname

    fade=ParseFade(ban,userid,password,companyname)
    fade.RegisterInvoiceToday()
    print ('Invoice Regist a day..',GetNow())

def RegisterMonth():
    global ban,userid,password,companyname

    fade=ParseFade(ban,userid,password,companyname)
    fade.RegisterInvoiceMonth()
    print ('Invoice Regist a month..',GetNow())

def RegisterDiscountAMonth():    
    global ban,userid,password,companyname

    fade=ParseFade(ban,userid,password,companyname)
    fade.RegisterDiscountMonth()
    print ('Discount Regist a month..',GetNow())

def DownloadInvoice():
    global ban,userid,password,companyname

    fade= ParseFade(ban,userid,password,companyname)
    fade.DownInvoice()
    print('Invoice Downloaded',GetNow())

def DownloadDiscount():   
    global ban,userid,password,companyname

    fade= ParseFade(ban,userid,password,companyname)
    fade.DownDiscount()
    print('Discount Downloaded',GetNow())
    

schedule.every(3).minutes.do(CheckTemp)
schedule.every(30).minutes.do(ClearMessageLog)

schedule.every().day.at('06:02').do(RegisterMonth)
schedule.every().day.at('08:13').do(DownloadInvoice)

schedule.every().day.at('08:33').do(RegisterDay)
schedule.every().day.at('10:33').do(DownloadInvoice)

schedule.every().day.at('10:43').do(RegisterDay)
schedule.every().day.at('12:43').do(DownloadInvoice)

schedule.every().day.at('13:00').do(RegisterDay)
schedule.every().day.at('15:10').do(DownloadInvoice)

schedule.every().day.at('15:33').do(RegisterDay)
schedule.every().day.at('17:33').do(DownloadInvoice)

#discount
schedule.every().day.at('06:35').do(RegisterDiscountAMonth)
schedule.every().day.at('08:35').do(DownloadDiscount)

schedule.every().day.at('10:34').do(RegisterDiscountAMonth)
schedule.every().day.at('12:34').do(DownloadDiscount)

schedule.every().day.at('14:31').do(RegisterDiscountAMonth)
schedule.every().day.at('16:31').do(DownloadDiscount)


if __name__=="__main__":
    
    while True:
        schedule.run_pending()
        time.sleep(1)
    
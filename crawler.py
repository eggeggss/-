# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup as bs4
import re
import datetime
import calendar

class ParseInvoiceWeb:

    def __init__(self,ban,userid,password,company):
        self.ban=ban
        self.userid=userid
        self.password=password
        self.compny=company
        self.loginurl="https://www.einvoice.nat.gov.tw/Login"
        self.token=self.GetToken()   
        self.getinvoiceurl= 'https://www.einvoice.nat.gov.tw/APB2BGVAN/Agency/AgencyInvoiceQuery!queryInvoiceJob'
        self.getdiscounturl='https://www.einvoice.nat.gov.tw/APB2BGVAN/Agency/AgencyAllowanceQuery!queryAllowanceJob'                       
        self.downloadurl="https://www.einvoice.nat.gov.tw/APB2BGVAN/Agency/AgencyInvoiceQuery!downloadJobExcel.shtml"
        self.downloaddiscounturl='https://www.einvoice.nat.gov.tw/APB2BGVAN/Agency/AgencyAllowanceQuery!downloadJobExcel.shtml'
        self.registerurl="https://www.einvoice.nat.gov.tw/APB2BGVAN/Agency/AgencyInvoiceQuery!addApply"
        self.registerdiscounturl="https://www.einvoice.nat.gov.tw/APB2BGVAN/Agency/AgencyAllowanceQuery!addApply"
        self.companyname=company
               
    #取得token
    def GetToken(self):
        token=""
        try:
            payload={
                'userType':'B',
                'loginType':'U',
                'loginWay':'W',
                'serviceType':'I',
                'serial':'',	
                'pincode':self.password,
                'signatur':'',	
                'ban':self.ban,
                'userID':self.userid,
                'password':self.password,
                'pid':self.userid,
                'orgType':'',	
                'bindata':'',	
                'typeCheck':'',	
                'radioB':'0',
                'textfield2':self.ban,
                'textfield3':self.userid,
                'textfield4':self.password,
                'textfield6':''}

            headers=requests.utils.default_headers()

            headers.update({
                'Host': 'www.einvoice.nat.gov.tw',
                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:57.0) Gecko/20100101 Firefox/57.0',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'zh-TW,zh;q=0.8,en-US;q=0.5,en;q=0.3',
                'Accept-Encoding': 'gzip, deflate, br',
                'Referer': 'https://www.einvoice.nat.gov.tw/',
                'Content-Type': 'application/x-www-form-urlencoded',
                'Content-Length': "282",
                'Cookie': '_ga=GA1.3.2003685348.1512790318; _gid=GA1.3.133122579.1512790318; JSESSIONID=vPpKhtQLv7W3KZnKNpgYHnBWwPfwRh8fz1CyNcvxmzGQpNwRl02J!-1853099029; PROXY_APMEMBERVAN=q3NGhtyS2Tv1DQ2CN4vVDThgtGfDH2XGh1vwNMsBwvqYMRDv1l4w!-635996108; TS01791bcd=01c54edd72e4bdfe7bbca21784edbb1c5f4f46b89dbed07569ae75dd10eed4abd84c43dabaf283df1d0255d1478d40517504d02f6af926e04a44723b595767e16549f160ef03028600e183ad82bf4435a8a4f23a2fdf7bf2f73baee0ee08f49377f323ba66e4137950cb57b9dc8f09f1a680354118a3ccd65912d2011d2608446b1c09450383462a24f54e1a79d47fcf957b3d26d3; PROXY_APB2BGVAN=CG5ThthCq4XytgqTT66fBcq3v9yQ6DMT1TGdgHLpJdXpnBhzGLlp!2008974032; _gat=1',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': "1",
                'Pragma': 'no-cache',
                'Cache-Control': 'no-cache',
                })

            s=requests.Session()
            rs=s.post(self.loginurl,headers=headers,data=payload)

            #print(rs.text)
            c=s.cookies.get_dict()
            token=c.get('JSESSIONID')
        except Exception as e:
            print ('ParseInvoiceWeb.GetToken fail:',e)
            token=''

        return token

    #找最新的下載檔的invoice no ,seq no
    def ToInvoicePage(self):

        return util.GetInfoRequest(self.getinvoiceurl,self.token)   

    #下載檔案
    def DownInvoice(self,seqno,invoice):
        
        if (seqno!='' and invoice!=''):
            util.DownloadRequest(self.downloadurl,self.token,seqno,invoice)

        print ('DownInvoice Ok...',util.getTime())
           
    #找最新的下載檔的invoice no ,seq no
    def ToDiscountPage(self):

        return util.GetInfoRequest(self.getdiscounturl,self.token)   
              
    #下載檔案
    def DownDiscount(self,seqno,invoice):

        if (seqno!='' and invoice!=''):

            util.DownloadRequest(self.downloaddiscounturl,self.token,seqno,invoice)

            print ('DownDiscount Ok...',util.getTime())
        
    def RegisterInvoiceMonth(self):

        util.RegisteraMonth(self.registerurl,self.token,self.companyname)

    def RegisterInvoicePerDay(self):
        util.RegisterPerday(self.registerurl,self.token,self.companyname)

    def RegisterDiscountMonth(self):
        util.RegisterDiscountaMonth(self.registerdiscounturl,self.token,self.companyname)

class util:

    @staticmethod
    def getTime():

         now=datetime.datetime.now()
         executetime=now.strftime(r"%Y-%m-%d %H:%M:%S")
         return executetime

    @staticmethod
    def getMonthPeriod():
        year=datetime.datetime.now().year-1911
        month=datetime.datetime.now().month
        day=datetime.datetime.now().day
        invoiceStartDate=''
        invoiceEndDate=''

        #當月最後一天
        today = datetime.date(int(datetime.datetime.now().year),month,day)
        lastday=calendar.monthrange(datetime.datetime.now().year,month)
        lastdate = today.replace(day=lastday[1])

        #當月第一天
        today = datetime.date.today()
        first = today.replace(day=1)

        if ((month % 2)==0):
            #雙數月抓2個月
            
            lastMonth = first - datetime.timedelta(days=1)
            
            myfirstday=lastMonth.replace(day=1)
            mylastdate=lastdate
            
            myfirstday_y=str(myfirstday.year-1911)
            myfirstday_m=myfirstday.month
            if len(str(myfirstday_m))==1:
                myfirstday_m='0'+str(myfirstday_m)
            myfirstdayst=str(myfirstday_y)+'/'+str(myfirstday_m)+'/01'

            mylastdate_y=str(mylastdate.year-1911)
            mylastdate_m=mylastdate.month
            mylastdate_d=mylastdate.day
            mylastdayst=str(mylastdate_y)+'/'+str(mylastdate_m)+'/'+str(mylastdate_d)
            invoiceStartDate=myfirstdayst
            invoiceEndDate=mylastdayst
            

        else:
            #單數月抓一個月
            myfirstday_y=str(first.year-1911)
            myfirstday_m=first.month
            if len(str(myfirstday_m))==1:
                myfirstday_m='0'+str(myfirstday_m)
            myfirstdayst=str(myfirstday_y)+'/'+str(myfirstday_m)+'/01'

            mylastdate_y=str(lastdate.year-1911)
            mylastdate_m=lastdate.month
            mylastdate_d=lastdate.day
            mylastdayst=str(mylastdate_y)+'/'+str(mylastdate_m)+'/'+str(mylastdate_d)
            invoiceStartDate=myfirstdayst
            invoiceEndDate=mylastdayst


        if len(str(day))==1:
            dayst='0'+str(day)
        else:
            dayst=str(day)

        invPeriod=str(year)+"/"+str(month)

        #print (invPeriod,invoiceStartDate,invoiceEndDate) 
        return invPeriod,invoiceStartDate,invoiceEndDate    

    @staticmethod
    def RegisterPerday(url,token,companyname):
        try:
            cookie = {
            '_ga': 'GA1.3.2003685348.1512790318',
            '_gid': 'GA1.3.133122579.1512790318',
            'JSESSIONID': token,
            'PROXY_APMEMBERVAN':'F5ShhtcJnTfknz22B9tPFH18BX648PvJYf21NNJy8cTX45CQTjGh!-635996108',
            'TS01791bcd': '01c54edd72ddf25c761d87f671db2c36b47a3e2ef9eecd6233879433954649d08f1eb262458c421ea175ac9c5c0341de504903867d3869c750fa57b90a0428c627f38c2b2c32b0754f570e50e1cdf8cbf424fb937930a5806dec939368649af8a52384db1d7ca70deaad9a958a72e16a0f87e708ac8a6655a153cd735e30d2f5ea8e601370!-635996108',
            }

            year=datetime.datetime.now().year-1911
            month=datetime.datetime.now().month
            day=datetime.datetime.now().day
            if len(str(day))==1:
                dayst='0'+str(day)
            else:
                dayst=str(day)

            querydate=str(year)+"/"+str(month)+"/"+dayst

            invoiceStartDate=querydate
            invoiceEndDate=querydate

            payload={
            'invPeriod':'',	
            'invoiceStartDate':invoiceStartDate,
            'invoiceEndDate':invoiceEndDate,
            'strCompanyBan':companyname,
            'invoiceStartNumber':'',	
            'invoiceEndNumber':'',
            'queryInvType':'0',
            'invType':'00',
            'extStatus':'0'}

            print ('invoiceStartDate:',invoiceStartDate,'invoiceEndDate:',invoiceEndDate)

            r=requests.post(url,cookies=cookie,data=payload)

        except Exception as e:
            print ('util.RegisterPerday fail',e)

    @staticmethod
    def RegisteraMonth(url,token,company):     
        try:
                cookie = {
                '_ga': 'GA1.3.2003685348.1512790318',
                '_gid': 'GA1.3.133122579.1512790318',
                'JSESSIONID': token,
                'PROXY_APMEMBERVAN':'F5ShhtcJnTfknz22B9tPFH18BX648PvJYf21NNJy8cTX45CQTjGh!-635996108',
                'TS01791bcd': '01c54edd72ddf25c761d87f671db2c36b47a3e2ef9eecd6233879433954649d08f1eb262458c421ea175ac9c5c0341de504903867d3869c750fa57b90a0428c627f38c2b2c32b0754f570e50e1cdf8cbf424fb937930a5806dec939368649af8a52384db1d7ca70deaad9a958a72e16a0f87e708ac8a6655a153cd735e30d2f5ea8e601370!-635996108',
                }
                
                invPeriod,invoiceStartDate,invoiceEndDate=util.getMonthPeriod()

                payload={
                'invPeriod':invPeriod,
                'invoiceStartDate':invoiceStartDate,#'106/11/01',
                'invoiceEndDate':invoiceEndDate,#'106/12/31',
                'strCompanyBan':company,
                'invoiceStartNumber':'',
                'invoiceEndNumber':'',
                'queryInvType':'0',
                'invType':'00',
                'extStatus':'0'}

                print ('invPeriod:',invPeriod,'invoiceStartDate:',invoiceStartDate,'invoiceEndDate:',invoiceEndDate)

                r=requests.post(url,cookies=cookie,data=payload)
        except Exception as e:
               print('Register a month fail:',e)

    @staticmethod
    def RegisterDiscountaMonth(url,token,company):
        try:        
            cookie = {
            '_ga': 'GA1.3.2003685348.1512790318',
            '_gid': 'GA1.3.133122579.1512790318',
            'JSESSIONID': token,
            'PROXY_APMEMBERVAN':'F5ShhtcJnTfknz22B9tPFH18BX648PvJYf21NNJy8cTX45CQTjGh!-635996108',
            'TS01791bcd': '01c54edd72ddf25c761d87f671db2c36b47a3e2ef9eecd6233879433954649d08f1eb262458c421ea175ac9c5c0341de504903867d3869c750fa57b90a0428c627f38c2b2c32b0754f570e50e1cdf8cbf424fb937930a5806dec939368649af8a52384db1d7ca70deaad9a958a72e16a0f87e708ac8a6655a153cd735e30d2f5ea8e601370!-635996108',
            }

            invPeriod,invoiceStartDate,invoiceEndDate=util.getMonthPeriod()

            payload={
            'invPeriod':invPeriod,
            'invoiceStartDate':invoiceStartDate,
            'invoiceEndDate':invoiceEndDate,
            'strCompanyBan':company,
            'allowanceStartNumber':'',
            'allowanceEndNumber':'',
            'invoiceStartNumber':'',
            'invoiceEndNumber':'',
            'queryInvType':	'0',
            'extStatus':'0'
            }

            print ('invPeriod:',invPeriod,'invoiceStartDate:',invoiceStartDate,'invoiceEndDate:',invoiceEndDate)
            r=requests.post(url,cookies=cookie,data=payload)

        except Exception as e:
               print('Register discount a month fail:',e)

    @staticmethod
    def DownloadRequest(url,token,seqno,invoice):
        try:    
            cookie = {
            '_ga': 'GA1.3.2003685348.1512790318',
            '_gid': 'GA1.3.133122579.1512790318',
            'JSESSIONID': token,
            'PROXY_APMEMBERVAN':'F5ShhtcJnTfknz22B9tPFH18BX648PvJYf21NNJy8cTX45CQTjGh!-635996108',
            'TS01791bcd': '01c54edd72ddf25c761d87f671db2c36b47a3e2ef9eecd6233879433954649d08f1eb262458c421ea175ac9c5c0341de504903867d3869c750fa57b90a0428c627f38c2b2c32b0754f570e50e1cdf8cbf424fb937930a5806dec939368649af8a52384db1d7ca70deaad9a958a72e16a0f87e708ac8a6655a153cd735e30d2f5ea8e601370!-635996108',
            }

            json={
                'seqNo':seqno,
                'agencyBan':invoice,
                'invoiceStartDate':'',	
                'invoiceEndDate':''
            }

            r=requests.post(url,cookies=cookie,data=json,stream=True)
            local_filename = url.split('/')[-1]

            with open(r'C:\Users\rogerroan\Downloads\22099478_ExcelReport.xls','wb') as f:
                for chunk in r.iter_content(chunk_size=4096): 
                        if chunk: # filter out keep-alive new chunks
                            f.write(chunk)
        except Exception as e:
            print ("DownloadRequest fail:",e)
    
    @staticmethod
    def GetInfoRequest(url,token):

        cookie = {
        '_ga': 'GA1.3.2003685348.1512790318',
        '_gid': 'GA1.3.133122579.1512790318',
        'JSESSIONID': token,
        'PROXY_APMEMBERVAN':'F5ShhtcJnTfknz22B9tPFH18BX648PvJYf21NNJy8cTX45CQTjGh!-635996108',
        'TS01791bcd': '01c54edd72ddf25c761d87f671db2c36b47a3e2ef9eecd6233879433954649d08f1eb262458c421ea175ac9c5c0341de504903867d3869c750fa57b90a0428c627f38c2b2c32b0754f570e50e1cdf8cbf424fb937930a5806dec939368649af8a52384db1d7ca70deaad9a958a72e16a0f87e708ac8a6655a153cd735e30d2f5ea8e601370!-635996108',
        }
     
        year=datetime.datetime.now().year-1911
        month=datetime.datetime.now().month
        day=datetime.datetime.now().day
        if len(str(day))==1:
            dayst='0'+str(day)
        else:
            dayst=str(day)

        querydate=str(year)+"/"+str(month)+"/"+dayst

        json={
            "queryApplyDateStart":querydate,
            "queryApplyDateEnd":querydate
        }
   
        r=requests.post(url,cookies=cookie,data=json)

        bs=bs4(r.text,'html.parser')

        tr=bs.find_all('form',{'method':'post'})
        #tr=bs.find_all('tr',{'class':'altrow'})

        #print (str(tr))

        invoice=''
        seqno=''

        for item in tr:
            if '處理完成' in str(item):
                bs=bs4(str(item),'html.parser')
                tds=bs.find_all('td',style='word-wrap:break-word;word-break:break-all')
                invoicepattern='[0-9]{8}'
                invoice=re.findall(invoicepattern,str(tds))[0]
                #print ("invoice:",invoice)

                tds2=bs4(str(tds),'html.parser')

                inputs=tds2.find_all('input',id='seqNo')[0]
                inputmatch='[0-9]{5}'
                seqno=re.findall(inputmatch,str(inputs))[0]

                #self.invoice=invoice
                #self.seqno=seqno
                print ("seqno:",seqno)
                break;
        
        return {'invoice':invoice,'seqno':seqno}




class ParseFade:
    def __init__(self,ban,userid,password,company):
       self.parseinvoiceweb=ParseInvoiceWeb(ban,userid,password,company)
       print ('ParseFade init')

    def DownInvoice(self):        
        dic=self.parseinvoiceweb.ToInvoicePage()
        self.parseinvoiceweb.DownInvoice(dic.get('seqno'),dic.get('invoice'))
        print ('ParseFade DownInvoice')

       
    def DownDiscount(self):       
        dic=self.parseinvoiceweb.ToDiscountPage()
        self.parseinvoiceweb.DownDiscount(dic.get('seqno'),dic.get('invoice'))
        print ('ParseFade DownDiscount')
 

    def RegisterInvoiceToday(self):       
        self.parseinvoiceweb.RegisterInvoicePerDay()
        print ('ParseFade RegisterInvoicePerDay')

    def RegisterInvoiceMonth(self):      
        self.parseinvoiceweb.RegisterInvoiceMonth()      
        print ('ParseFade RegisterInvoiceMonth')

    def RegisterDiscountMonth(self):      
        self.parseinvoiceweb.RegisterDiscountMonth()
        print ('ParseFade RegisterDiscountMonth')


if __name__=="__main__":
    ban=''
    userid=''
    password=''
    company=''
    fade= ParseFade(ban,userid,password,company)
    #fade.RegisterInvoiceToday()
    #fade.RegisterInvoiceMonth()
    fade.RegisterDiscountMonth()
    #fade.DownInvoice()
    #fade.DownDiscount()
else:
    pass
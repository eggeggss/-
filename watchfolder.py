# -*- coding: utf-8 -*-
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler
import time,datetime
import subprocess
import os,sys
import logging
import xlrd
import csv
import pymssql

class MyInvoiceDetail:

    def __init__(self,invoice_no,dt_invoice,serial_no,
        description,qty,unit,price,total,qty2,unit2,price2,total2,comment):
            
        self.invoice_no=invoice_no
        self.dt_invoice=dt_invoice
        self.serial_no=serial_no
        self.description=description
        self.qty =qty
        self.unit=unit
        self.price =price
        self.total =total
        self.qty2 =qty2
        self.unit2 =unit2
        self.price2 =price2
        self.total2 =total2
        self.comment =comment

class MyInvoice:
    def __init__(self,invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,cost2,costtype):
        self.invoice_no=invoice_no
        self.comment=comment
        self.f_type=f_type
        self.f_status=f_status
        self.dt_invoice=dt_invoice
        self.buyerno=buyerno
        self.buyername=buyername
        self.sellerno=sellerno
        self.sellername=sellername
        self.dt_seller=dt_seller
        self.cost=cost
        self.cost1=cost1
        self.cost2=cost2
        self.costtype=costtype

class Discount:
    def __init__(self,discount_no ,format_type  ,status  ,category   ,
		invoice_no  ,dt_invoice  ,buyer_no ,
		buyer_name  ,seller_no  ,seller_name  ,
		dt_seller  ,item_name   ,itemdiscount_notax  ,
		itemdiscount_tax   ,discount_notax ,discount_tax   ,
		comment  ,dt_discount):

        self.discount_no=discount_no
        self.format_type=format_type
        self.status=status
        self.category=category
        self.invoice_no=invoice_no
        self.dt_invoice=dt_invoice
        self.buyer_no=buyer_no
        self.buyer_name=buyer_name
        self.seller_no=seller_no
        self.seller_name=seller_name
        self.dt_seller=dt_seller
        self.item_name=item_name
        self.itemdiscount_notax=itemdiscount_notax
        self.itemdiscount_tax=itemdiscount_tax
        self.discount_notax=discount_notax
        self.discount_tax=discount_tax
        self.comment=comment
        self.dt_discount=dt_discount


class DAOw:

    def __init__(self):
       pass
    
    def GetConnect(self):
        return pymssql.connect(host='0.0.0.0', user='id',password='pass',database='db')

    def InsertInvoiceS(self,invoices):
        try:
            conn=self.GetConnect()
            cursor=conn.cursor()
            #cursor.execute("insert into MIRLE_APP.dbo.invoice(invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,total,costtype,dt_create,stat_void) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',{10},{11},{12},'{13}',getdate(),0)".format(invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,cost2,costtype))
            #cursor.execute("exec MIRLE_APP.dbo.zp_insertinvoice '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}'".format(invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,cost2,costtype))

            for invoice in invoices:
                cursor.execute("exec MIRLE_APP.dbo.zp_insertinvoice '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}'".format(invoice.invoice_no,
                invoice.comment,
                invoice.f_type,
                invoice.f_status,
                invoice.dt_invoice,
                invoice.buyerno,
                invoice.buyername,
                invoice.sellerno,
                invoice.sellername,
                invoice.dt_seller,
                invoice.cost,
                invoice.cost1,
                invoice.cost2,
                invoice.costtype))
                print ('write master:',invoice.invoice_no)

            conn.commit()

            conn.close()
        except Exception as e:
            print ("DAOw InsertInvoiceS fail:",e)


    def InsertInvoice(self,invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,cost2,costtype):
        try:
            conn=self.GetConnect()
            cursor=conn.cursor()
            #cursor.execute("insert into MIRLE_APP.dbo.invoice(invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,total,costtype,dt_create,stat_void) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',{10},{11},{12},'{13}',getdate(),0)".format(invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,cost2,costtype))
            cursor.execute("exec MIRLE_APP.dbo.zp_insertinvoice '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}'".format(invoice_no,comment,f_type,f_status,dt_invoice,buyerno,buyername,sellerno,sellername,dt_seller,cost,cost1,cost2,costtype))

            conn.commit()

            conn.close()
        except Exception as e:
            print ("DAOw fail:",e)

    def InsertInvoiceDetail(self,details):
        try:
            conn=self.GetConnect()
            cursor=conn.cursor()
            
            for detail in details:
                cursor.execute("exec MIRLE_APP.dbo.zp_insertinvoicedetail '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}'"
                .format(
                detail.invoice_no,
                detail.dt_invoice,
                detail.serial_no,
                detail.description,
                detail.qty,
                detail.unit,
                detail.price,
                detail.total,
                detail.qty2,
                detail.unit2,
                detail.price2,
                detail.total2,
                detail.comment
                ))
                print ('write detail:',detail.invoice_no)

            conn.commit()

            conn.close()
        except Exception as e:
            print ("DAOw InsertInvoiceDetail fail:",e)
    
    def InsertDiscount(self,discounts):
        try:
            conn=self.GetConnect()
            cursor=conn.cursor()
            
            for discount in discounts:
                cursor.execute("exec MIRLE_APP.dbo.zp_insertdiscount '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}', '{11}','{12}','{13}','{14}','{15}','{16}','{17}'"
                .format(
                discount.discount_no ,discount.format_type  ,discount.status  ,discount.category   ,
		        discount.invoice_no  ,discount.dt_invoice  ,discount.buyer_no ,
		        discount.buyer_name  ,discount.seller_no  ,discount.seller_name  ,
		        discount.dt_seller  ,discount.item_name   ,discount.itemdiscount_notax  ,
		        discount.itemdiscount_tax   ,discount.discount_notax ,discount.discount_tax   ,
		        discount.comment  ,discount.dt_discount
                ))
                print ('write discount:',discount.discount_no)

            conn.commit()

            conn.close()
        except Exception as e:
            print ("DAOw InsertDiscount fail:",e)




class  MyEH(LoggingEventHandler):
    #production   lindafan 
    #reciever="rogerroan@mirle.com.tw;lindafan@mirle.com.tw"
    #test
    reciever="rogerroan@mirle.com.tw"

    def __init__(self):
        super(MyEH, self).__init__()
        

    def on_created(self , event):
        self.process(event)

    def on_modified(self , event):
        self.process(event)
        
    def process(self , event):
        global src,des
        if(event.event_type=='created' or event.event_type=='modified'):
            
            name,extension=os.path.splitext(event.src_path)
            
            if (extension==".xls"):

                #取得檔案名稱
                filename=os.path.basename( os.path.abspath(event.src_path))
                
                if (filename=='22099478_ExcelReport.xls'):
                    #檔名為日期時間
                    now=datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
                    filename=now+".xls"
                    xlsname=now+".csv"
                    #join name
                    destname=os.path.join(des,filename)
                    csvdestname=os.path.join(des,xlsname)

                    command="move "+event.src_path+" "+destname
                    #搬檔案
                    os.system("move "+event.src_path+" "+destname)

                    #insert master
                    to= self.csv_from_excelmaster(destname,csvdestname)  

                    if to!='':
                        self.insertdbmaster(to)
    
                    #insert detail

                    xlsnamedetail=now+"_detail"+".csv"
                    csvdestnamedetail=os.path.join(des,xlsnamedetail)
                    to1= self.csv_from_exceldetail(destname,csvdestnamedetail)

                    if to1!='':
                        self.insertdbdetail(to1)


                    xlsnamediscount=now+"_discount"+".csv"
                    csvdestnamediscount=os.path.join(des,xlsnamediscount)
                    to2= self.csv_from_exceldiscount(destname,csvdestnamediscount)

                    if to2!='':
                        self.insertdbdiscount(to2)

                    #csv_from_exceldiscount
                    #print('send',destname,csvdestname)
                    
                    subject="盟立發票紀錄 時間:"+datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.send_mail("rogerroan@mirle.com.tw",self.reciever,subject,
                    "","192.168.2.7",25,"","",False) 
                

    def csv_from_excelmaster(self,fromfile,tofile):
        try:
            wb = xlrd.open_workbook(fromfile)
       
            if '折讓單' in wb.sheet_names():
                return ""

            sheetname="" 
            if '發票主檔' in  wb.sheet_names():
                sheetname='發票主檔'
            else:
                sheetname='Sheet0'

            sh = wb.sheet_by_name(sheetname)
            your_csv_file = open(tofile, 'w', encoding='utf8')
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

            for rownum in range(sh.nrows):
                wr.writerow(sh.row_values(rownum))

            your_csv_file.close()
            return tofile
        
        except Exception as e:
            print ("csv_from_excelmaster fail:",e)
            return ""


    def csv_from_exceldetail(self,fromfile,tofile):
        try:
            wb = xlrd.open_workbook(fromfile)
       
            if '折讓單' in wb.sheet_names():
                return ""

            sheetname="" 
            if '發票明細' in  wb.sheet_names():
                sheetname='發票明細'
            else:
                sheetname='Sheet1'

            sh = wb.sheet_by_name(sheetname)
            your_csv_file = open(tofile, 'w', encoding='utf8')
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

            for rownum in range(sh.nrows):
                wr.writerow(sh.row_values(rownum))

            your_csv_file.close()
            return tofile
        
        except Exception as e:
            print ("csv_from_excelmaster fail:",e)
            return ""

    def csv_from_exceldiscount(self,fromfile,tofile):
        try:
            wb = xlrd.open_workbook(fromfile)

            if '折讓單' not in  wb.sheet_names():
                return ""

            sheetname="" 
            if '折讓單' in  wb.sheet_names():
                sheetname='折讓單'           

            sh = wb.sheet_by_name(sheetname)
            your_csv_file = open(tofile, 'w', encoding='utf8')
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

            for rownum in range(sh.nrows):
                wr.writerow(sh.row_values(rownum))

            your_csv_file.close()
            return tofile
        
        except Exception as e:
            print ("csv_from_exceldiscount fail:",e)
            return ""


    def insertdbmaster(self,tofile):
         
        #f = open(r'D:\pythonproject\invoice\excel\2017_12_04_17_36_40.csv', encoding='utf8')
        f = open(tofile, encoding='utf8')
        
        mylist=list()

        for row in csv.DictReader(f):
                invoice=MyInvoice(
                row['發票號碼'],
                row['註記欄(不轉入進銷項媒體申報檔)'],
                row['格式代號'],
                row['發票狀態'], 
                row['發票日期'],
                row['買方統一編號'],
                row['買方名稱'],
                row['賣方統一編號'],
                row['賣方名稱'],
                row['寄送日期']  , 
                row['銷售額合計'],
                row['營業稅'],
                row['總計'],
                row['課稅別'])

                mylist.append(invoice)

              
        f.close()

        dal=DAOw()
        dal.InsertInvoiceS(mylist)
        print ('master 寫入完成')
    

    def insertdbdetail(self,tofile):
         
        f = open(tofile, encoding='utf8')   
        mylist=list()
        for row in csv.DictReader(f):
                invoice=MyInvoiceDetail(
                row['發票號碼'],
                row['發票日期'],
                row['序號'],
                row['發票品名'].replace("'",""), 
                row['數量'],
                row['單位'],
                row['單價'],
                row['小計'],
                row['數量2(鋼鐵業專用)'],
                row['單位2(鋼鐵業專用)']  , 
                row['單價2(鋼鐵業專用)'],
                row['小計2'],
                row['單一欄位備註'],
                )
                mylist.append(invoice)

        f.close()
        dal=DAOw()
        dal.InsertInvoiceDetail(mylist)
        print ('detail 寫入完成')

    def insertdbdiscount(self,tofile):
        
        f = open(tofile, encoding='utf8')   
        mylist=list()
        
        for row in csv.DictReader(f):
            discount=Discount(
                row['折讓單號碼'],
                row['格式代號'],
                row['折讓單狀態'],
                row['折讓單類別'], 
                row['發票號碼'],
                row['發票日期'],
                row['買方統一編號'],
                row['買方名稱'],
                row['賣方統一編號'],
                row['賣方名稱']  , 
                row['寄送日期'],
                row['品項名稱'],
                row['品項折讓金額(不含稅)'],
                row['品項折讓稅額']  , 
                row['折讓金額(不含稅)'],
                row['折讓稅額'],
                row['註記欄(不轉入進銷項媒體申報檔)'],
                row['折讓單日期'],
            )

            mylist.append(discount)

        f.close()
        dal=DAOw()
        dal.InsertDiscount(mylist)
        print ('discount 寫入完成')


    def send_mail(self, send_from,send_to,subject,text,server,port,username='',password='',isTls=True,filepath='',filename=''):
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = send_to
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        
        '''
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(filepath, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(filename))
        msg.attach(part)
        '''
        #context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
        #SSL connection only working on Python 3+
        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        #smtp.login(username,password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()      

src=r'''C:\Users\rogerroan\Downloads'''
des=r'''D:\pythonproject\invoice\excel'''
 


if __name__=="__main__":

    #fade=ParseFade()
    #fade.DownDiscount()
    #dao=DAOw()

    #dao.InsertInvoice('invoice_no2',	'comment2',	'f_type',	'f_status',	'dt_invoice',	'buyerno',	'buyername','sellerno',	'sellername',	'dt_seller',	0,	1,	2,	'costtype')
 
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S')
 
    event_handler = MyEH()
    observer = Observer()
    observer.schedule(event_handler, src, recursive=True)
    observer.start()
    print("listen invoice download")
    while(True):
         time.sleep(1)
    
    
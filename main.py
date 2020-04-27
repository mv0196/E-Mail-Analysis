from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
from PyQt5.QtChart import QChart, QChartView, QValueAxis, QBarCategoryAxis, QBarSet, QBarSeries
from PyQt5.Qt import Qt
from PyQt5.QtGui import QPainter

import pandas as pd
import numpy as np
import matplotlib.pyplot as pp
import re

import mailbox
import csv
import email
import imaplib
import getpass, poplib
import warnings
from os import getcwd

warnings.filterwarnings("ignore",category=DeprecationWarning)

login_ui,_=loadUiType('login.ui')
index_ui,_=loadUiType('index.ui')
graph_ui,_=loadUiType('graph_dialog.ui')

path_to_file=getcwd()+"\\files\\"


class Login(QMainWindow,login_ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.password.setEchoMode(QLineEdit.Password)
        self.showpass.stateChanged.connect(self.showpassword)
        self.login_btn.clicked.connect(self.handle_login)

    def showpassword(self,checked):
        if checked:
            self.password.setEchoMode(QLineEdit.Normal)
        else:
            self.password.setEchoMode(QLineEdit.Password)
    def handle_login(self):
        self.login_btn.setEnabled(False)
        username=self.user.text().split("@")[0]
        password=self.password.text()

        msg=QMessageBox()
        if username=="":
            msg.setText("Email Id can't be empty!")
            msg.setIcon(QMessageBox.Warning)
            x=msg.exec_()
            self.login_btn.setEnabled(True)
            return
        if password=="":
            msg.setText("Password can't be empty!")
            msg.setIcon(QMessageBox.Warning)
            x=msg.exec_()
            self.login_btn.setEnabled(True)
            return

        self.setEnabled(False)
        c = imaplib.IMAP4_SSL("imap.gmail.com")
        try:
            c.login(username,password)
            c.select("inbox")
            #no. of emails to b analysed
            n,done=QtWidgets.QInputDialog.getInt(self,"No. of mails",'Enter the number of emails you want to analyse:')
            n=abs(n)#getting absolute value of n in case if user entered -ve value
            msg.setText("Please wait for sometime!,We are fetching data.")
            msg.setIcon(QMessageBox.Information)
            msg.setInformativeText("Click ok to Start and wait")
            x=msg.exec_()
            with open (path_to_file+username+".csv",'w',newline='') as o:
                writer=csv.writer(o)
                writer.writerow([ 'subject', 'to', 'from' ,'Date'])
                for i in range(1,n):
                    typ, msg_data = c.fetch(str(i), '(RFC822)')
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            message = email.message_from_bytes(response_part[1])
                        # for header in [ 'subject', 'to', 'from' ,'Date']:
                            # print ('%-8s: %s' % (header.upper(), msg[header]))
                            # print (msg[header])

                            writer.writerow([message['subject'],message['to'],message['from'],message['Date']])
            msg.setText("File Writen Succesfully")
            msg.setInformativeText("")
            msg.setDetailedText("")
            x=msg.exec_()
            self.index=Index(username)
            self.close()
            self.index.show()

        except:
            msg.setText('Wrong UserName or Password! or You have not provided access to,less secure app from your google account for fetching mails')
            msg.setIcon(QMessageBox.Critical)
            msg.setDetailedText("To allow less secure apps visit https://myaccount.google.com/lesssecureapps?utm_source=google-account&utm_medium=web")
            msg.setInformativeText("Check details and try again")
            x=msg.exec_()
            self.setEnabled(True)
            self.login_btn.setEnabled(True)

class Index(QMainWindow,index_ui):
    def __init__(self,username):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.radioButton1.toggled.connect(lambda:self.radiobutton1(self.radioButton1))
        self.radioButton2.toggled.connect(lambda:self.radiobutton2(self.radioButton2))
        self.radioButton3.toggled.connect(lambda:self.radiobutton3(self.radioButton3))
        self.radioButton4.toggled.connect(lambda:self.radiobutton4(self.radioButton4))
        self.radioButton5.toggled.connect(lambda:self.radiobutton5(self.radioButton5))
        self.radioButton6.toggled.connect(lambda:self.radiobutton6(self.radioButton6))
        self.radioButton7.toggled.connect(lambda:self.radiobutton7(self.radioButton7))
        self.radioButton8.toggled.connect(lambda:self.radiobutton8(self.radioButton8))
        self.radioButton9.toggled.connect(lambda:self.radiobutton9(self.radioButton9))

        self.df=pd.read_csv(path_to_file+username+'.csv')#path to store and retrieve files
        self.df['from']=self.df['from'].dropna().apply(self.clean)#clean is fuction made above
        self.df['to']=self.df['to'].dropna().apply(self.clean)
        self.df['Date']=self.df['Date'].apply(lambda s:pd.to_datetime(s).tz_convert('Asia/Kolkata'))

    def clean(self,raw):
        match=re.search('<(.+)>',raw)
        if match is None:
            return raw
        else:
            return match.group(1)


    def radiobutton1(self,btn):
        if btn.isChecked():
            n,done=QtWidgets.QInputDialog.getInt(self,"Top Mailers","Enter the number of Top Mailers you want to know about(Your max mailers are- "+str(len(list(set(self.df['from']))))+"):")
            n=abs(n)
            self.top_mailers(n)
    def radiobutton2(self,btn):
        if btn.isChecked():
            n,done=QtWidgets.QInputDialog.getInt(self,"Top Dates","Enter the number of Top Dates you want to know about:",)#(Your max ----- are- ",len(list(set(self.df['dateofmonth']))),")   yha pe ye isliye ni lgaya kuki dateofmonth wala coloumn function k andr bna h!,use phle acess ni kr skte"
            n=abs(n)
            self.top_dates(n)
    def radiobutton3(self,btn):
        if btn.isChecked():
            self.top_days_of_week()
    def radiobutton4(self,btn):
        if btn.isChecked():
            n,done=QtWidgets.QInputDialog.getInt(self,"Top Months","Enter the number of Top months you want to know about:",)
            n=abs(n)
            self.top_months(n)
    def radiobutton5(self,btn):
        if btn.isChecked():
            n,done=QtWidgets.QInputDialog.getInt(self,"Top Time Range","Enter the number of Top Time Range you want to know about:",)
            n=abs(n)
            self.top_time_range(n)
    def radiobutton6(self,btn):
        if btn.isChecked():
            n,done=QtWidgets.QInputDialog.getInt(self,"Time Top Mailers Mail","Enter the number of Top Mailers you want to know about the Most time they mail:",)
            n=abs(n)
            self.time_top_mailers_mail(n)
    def radiobutton7(self,btn):
        if btn.isChecked():
            n,done=QtWidgets.QInputDialog.getInt(self,"Top Mailers Time Range","Enter the number of Top Mailers you want to know about:",)
            n=abs(n)
            self.top_mailers_time_range(n)
    def radiobutton8(self,btn):
        if btn.isChecked():
            n,done=QtWidgets.QInputDialog.getInt(self,"Top Year","Enter the number of Top Years you want to know about:",)
            n=abs(n)
            self.top_year(n)
    def radiobutton9(self,btn):
        if btn.isChecked():
            self.max_mail_date()

    def plotbargraph(self,x,name_of_chart):#x will ber series object
        set0=QBarSet('X0')#It will contain value for each x axis
        set0.append(list(x.values))#x.values will ne numpy nd array

        series=QBarSeries()
        series.append(set0)

        chart=QChart()
        chart.addSeries(series)
        chart.setTitle(name_of_chart)
        chart.setAnimationOptions(QChart.SeriesAnimations)

        axisX=QBarCategoryAxis()
        axisX.append(list(map(str,x.axes[0])))#x.axes[0] will give list of axes

        axisY=QValueAxis()
        axisY.setRange(0,max(list(x.values)))

        chart.addAxis(axisX,Qt.AlignBottom)
        chart.addAxis(axisY,Qt.AlignLeft)

        chart.legend().setVisible(False)#Set True to show legend
        #chart.legend().setAlignment(Qt.AlignBottom)

        chartView=QChartView(chart)
        self.display_chart=GraphDialog(chartView)
        self.display_chart.show()

    def top_mailers(self,n):
        try:
            #for top n mailers
            countm=self.df['from'].value_counts()
            #print(countm)
            #type(countm)
            self.plotbargraph(countm.head(n),"Top Mailers")
        except:
            pass
    def top_dates(self,n):
        try:
            #for number of mails on each date
            self.df['dateofmonth']=self.df['Date'].dt.day#day gives day of months bt dayofyear gives days out of 365
            #self.df['dateofmonth']
            countdt=self.df['dateofmonth'].value_counts()
            if n>max(countdt.axes[0]):#is is fer..lets say in data we have mails ony for 3 dates but user entered 5 top dates
                n=max(countdt.axes[0])
            self.plotbargraph(countdt.head(n),"Top Dates")
        except:
            pass

    # In[9]:


    def top_days_of_week(self):
        try:
            #for number of mails on each day of week
            self.df['dayofweek']=pd.Categorical(self.df['Date'].dt.day_name(),ordered=True,categories=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'])
            #pd.Categorical allows values only within a category,i.e. actegories mentioned
            #self.df['dayofweek']
            countdw=self.df['dayofweek'].value_counts()
            self.plotbargraph(countdw,"Top Days od Week")
        except:
            pass

    def top_months(self,n):
        try:
            #for number of mails in each month
            self.df['monthofyear']=self.df['Date'].dt.strftime('%b')#month give number of munth i.e. for jan-1,for feb-2,etc & month_name gives ajeeb o/p
            #strftime('%b')    --->convert into months name
            #self.df['monthofyear']
            countmy=self.df['monthofyear'].value_counts()
            if n>len(countmy.axes[0]):
                n=len(countmy.axes[0])
            self.plotbargraph(countmy.head(n),"Top Months")
        except:
            pass

    def top_time_range(self,n):
        try:
            # for number of mails in each time range
            self.df['timeofday']=self.df['Date'].dt.hour+self.df['Date'].dt.minute/60
            #type(self.df['timeofday'][0])   ----->numpy.float64   to find range of time convert it into float
            self.df['timeofday(hour_only)']=self.df['Date'].dt.hour
            countt=self.df['timeofday(hour_only)'].value_counts()
            if n>len(countt.axes[0]):
                n=len(countt.axes[0])
            self.plotbargraph(countt.head(n),"Top Time Range")
        except:
            pass

    def time_top_mailers_mail(self,n):
        try:
            #to find out when mostly the mail of top n mailers come
            l=[]
            #
            countm=self.df['from'].value_counts()   #for top mailers
            for i in countm.head(n).axes[0]:
                #print(i)
                count=self.df[self.df['from']==i]['timeofday(hour_only)'].value_counts()
                #print(count)
                #print('*')
                l.append(count.axes[0][0])#appending top time range of each mailer of countm.head(5).axes[0] in l,count.head(1) will give number of emails but the hours are at axes and axes[0 will give list of all axes for timeofday(hour_only) for each i
            #print(l)
            l=pd.Series(l,index=countm.head(n).axes[0])
            self.plotbargraph(l,"When Top Mailers Mail")
        except:
            pass

    # In[13]:


    def top_mailers_time_range(self,n):
        try:
            l=[]
            #for top m mailers
            countm=self.df['from'].value_counts()
            #countm.head(m)
            a,done=QtWidgets.QInputDialog.getInt(self,"Start Time","Enter the start of time range(Time in 24hrs standard):")
            b,done=QtWidgets.QInputDialog.getInt(self,"End Time","Enter the end of time range:")
            a=abs(a)
            b=abs(b)
            self.df['timeofday(hour_only)']=self.df['Date'].dt.hour
            for i in countm.head(n).axes[0]:#axes gives pd series but axes[0] gives list of all axes and axes[1] will be out of range
                #print(i)    #must try
                nl=[]
                for i in self.df[self.df['from']==i]['timeofday(hour_only)']:
                    #print("*"+str(i)+"*")
                    if i>=a and i<b:
                        nl.append(i)
                #print(nl)    #must try
                count=pd.Series(nl).value_counts()
                #print(list(count.axes))    #must try
                #axis of count contain time and value is number of times mail came at that time
                #print(count)
                #print('*')
                #print(count.axes[0][0])#gives time at which most mails come,bcoz count.axes[0] will give lsit of all indexes
                try:
                    l.append(count.values[0])#acessing top value i.e. most number of mails,axes of nl have the hours at which mail arrived
                    #to find number of mails arrived instead of "count.axes[0][0]" apeend "count[1]" to l
                except:
                    l.append(0)#i.e no mail from this mailer in this time

            l=pd.Series(l,index=countm.head(n).axes[0])
            self.plotbargraph(l,"Top Mailers Time Range")

        except:
            pass

    # In[14]:


    def top_year(self,n):
        try:
            self.df['year']=self.df['Date'].dt.year
            county=self.df['year'].value_counts()
            if n>len(county.axes[0]):
                n=len(county.axes[0])
            self.plotbargraph(county.head(n),"Top Year")
        except:
            pass

    # In[15]:


    def max_mail_date(self):
        try:
            y=self.df['Date'].dt.year.value_counts().axes[0][0]#gives year with maximum mails
            m=self.df[self.df['Date'].dt.year==y]['Date'].dt.month.value_counts().axes[0][0]#gives month number in whic max mails arrived in year 'y'
            #print(m)
            d=self.df[self.df['Date'].dt.month==m]['Date'].dt.day.value_counts().axes[0][0]#gives date on ehic max mails arrived in month 'm' in year 'y'
            #print(d)
            msg=QMessageBox()
            msg.setText('Maximum Mails arrived on:(dd/mm/yyyy)'+str(d)+'-'+str(m)+'-'+str(y))
            x=msg.exec_()
        except:
            pass


class GraphDialog(QMainWindow,graph_ui):
    def __init__(self,chartView):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.setCentralWidget(chartView)#chartdisplay is name of widget in graph_ui


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window=Login()
    window.show()
    sys.exit(app.exec_())

from tkinter import *
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import yfinance as yf
import datetime as dt
from datetime import datetime, date
import xlsxwriter
import openpyxl

def disappear(a):
    a.place(x=0,y=0,width=0,height=0)

def OK():
    disappear(entry1)
    disappear(button)
    label1.config(text='The file has been updated')
    label1.place(x=150, y=80, width=200, height=20)
    global name
    name=entry1.get()
    determining_dates()
    analysis()

def determining_dates():
    global end_date
    global start_date
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.strptime(end_date, '%Y-%m-%d') - dt.timedelta(days=365)).strftime('%Y-%m-%d')

def analysis():
    global name
    stock = yf.Ticker(name)
    msci = yf.Ticker('msci')
    global end_date
    global start_date
    print(end_date)
    print(start_date)

    stock_hist = stock.history(start='2023-10-1', end=end_date)
    stock_data= pd.DataFrame(stock_hist)
    stock_data['daily return %'] = round((stock_data['Close'] - stock_data['Open']) / stock_data['Open'] * 100,2)
    global msci_data
    msci_hist= msci.history(start='2023-10-1', end=end_date)
    msci_data= pd.DataFrame(msci_hist)
    msci_data['daily return %'] = round((msci_data['Close'] - msci_data['Open']) / msci_data['Open'] * 100,2)

    x = msci_data['daily return %'].tolist()

    y = stock_data['daily return %'].tolist()

    a, b = linear_regression(x, y)
    global data
    data=pd.DataFrame(stock_data['daily return %'].describe())
    average=data.at['mean','daily return %']
    standard_deviation = data.at['std', 'daily return %']
    minimum = data.at['min', 'daily return %']
    maximum = data.at['max', 'daily return %']
    data=[name,average,standard_deviation,minimum,maximum,b]
    saving()
    plotting(stock_data,'ημερήσια απόδοση',msci_data['daily return %'].tolist())
    plotting(stock_data, 'ημερήσια τιμή κλεισίματος',msci_data['Close'].tolist())


def saving():
    global data
    global path
    df=pd.read_excel(path+'Αρχείο Μετοχών.xlsx',header=None)
    print(df)
    df[name]=data
    csvlist='Μετοχή','Μέση ημερήσια απόδοση(%)','Τυπική Απόκλιση','Ελάχιστη ημερήσια απόδοση(%)','Μέγιστη ημερήσια απόδοση(%)','Συντελεστής β'
    csvdf=pd.DataFrame(np.array(csvlist).reshape((6,1)))
    csvdf[name]=df[name]
    csvdf.to_csv(path +name+ ' analysis.txt', sep=' ', header=False,index=False)
    print(df)
    writer = pd.ExcelWriter(path + 'Αρχείο Μετοχών.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False,header=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    for i, col in enumerate(df.columns):
        width = max(df[col].apply(lambda x: len(str(x))).max(), len(df[col]))
        worksheet.set_column(i, i, width)
    writer.close()


def plotting(df,k,msci):
    global name
    global msci_data
    if k== 'ημερήσια απόδοση':
        l = 'ημερήσιες αποδόσεις'
        a='daily return %'
    else:
        l = 'ημερήσιες τιμές κλεισίματος'
        a='Close'

    df.index=x = [*(range(0, len(df)))]

    x = [*(range(0, len(df)))]
    y = df[a].tolist()
    print(len(x), len(y))

    figure1= plt.figure(figsize=(18, 10))
    x = [*(range(10, len(df)))]
    list = []
    for i in x:
        list.append(kiliomenos_mesos(df, a, i, 10))
    plt.plot(x, list, c='r')
    plt.ylabel(k + '(%)')
    plt.xlabel('χρόνος')
    plt.title(
        'διαχρονική εξέλιξη για την μέση\n ' + k + '(με κυλιόμενο μέσο που\n αξιοποιεί τις δέκα τελευταίες μετρήσεις)')
    plt.savefig(path +name+' '+ l + ' με κυλιόμενο μέσο(δίαγραμμα).png')

    figure2 = plt.figure(figsize=(18, 10))
    x = [*range(0, len(df))]
    plt.xticks([])
    plt.plot(x, msci, c='r')
    plt.plot(x, df[a], c='b')
    plt.xlabel('χρόνος')
    plt.ylabel(k + '(%)')
    plt.title('με μπλε η μετοχή\n με κόκκινο ο msci')
    plt.savefig(path + name+' '+l + ' μετοχής και δείκτη(διάγραμμα).png')




def linear_regression(x,y):
    xvar = float(sum(x) / len(x))
    yvar = sum(y) / len(y)
    b1 = 0
    b2 = 0
    for i in range(0, len(y)):
        b1 += (x[i] - xvar) * (y[i] - yvar)
        b2 += (x[i] - xvar) ** 2
    b = b1 / b2
    a = yvar - b * xvar
    print(a,b)
    a=round(a,4)
    b=round(b,4)
    return[a,b]

def kiliomenos_mesos(df,k,place,n):
    rangee=range(place-10,place+1)
    sum=0
    for i in rangee:
        sum+=df.at[i,k]

    mesos=sum/(n+1)
    return(mesos)

window=Tk()
window.geometry('500x500')
window.title('Stock Analysis')

label1=Label(window,text='Name of the Stock: ')
entry1=Entry(window)
button=Button(window,text='OK',command=OK)

label1.place(x=20,y=80,width=200,height=20)
entry1.place(x=230,y=80,width=200,height=20)
button.place(x=430,y=140)

global path
path=""





window.mainloop()
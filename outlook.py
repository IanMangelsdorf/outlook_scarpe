import win32com.client
import time
import datetime as dt
from os import path
import pandas as pd
import re

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy import exists
from sqlalchemy import exists, or_, and_

from sql import Emails,Base
import pywintypes

import multiprocessing


engine = create_engine('sqlite:///outlook.db')

Base.metadata.bind = engine

DBSession = sessionmaker(bind=engine)
session = DBSession()


from win32com.client.gencache import EnsureDispatch as Dispatch

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
Mainfolder=''

df = pd.DataFrame([[0,0,0,0,0,0,0,0,0]],columns=['Seneder', 'Subject', 'Date','Sender Email',
                           'Company','MainFolder','Folder','Recipient', 'Phone'])




phone_patten = '(\(+61\)|\+61|\(0[1-9]\)|0[1-9])?( ?-?[0-9]){6,9}'

class Oli():
    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        array_size = self._obj.Count
        for item_index in range(1,array_size+1):
            yield (item_index, self._obj[item_index])

    def prop(self):
        return sorted( self._obj._prop_map_get_.keys() )


def company(senderEmail):
    if "EXCHANGELABS" in senderEmail or 'O=ECSM' in senderEmail :
        email = "Internal"
        comp = "Internal"
    else:
        x = senderEmail.split("@")[-1]
        comp = x.split('.')[0]
        email = senderEmail

    return email, comp

def phone (body):
    ph1 = "Unknown"
    phone_types =['t |', 'm |', 'phone', 'mobile', 'cell' ]
    n = body.splitlines()
    for inx,i in enumerate(n[::-1]):
        if 'regards' in i.lower():
            lst = n[(len(n)-inx):]
            lst = [string for string in lst if string != ""]
            for lin in lst:
                for typ in phone_types:
                    if typ.lower() in lin.lower():
                        return lin

    return ph1

def recipients(recip):
    lst =[]
    for r in recip:
        lst.append(r.Name)
    lst.append('Unknown')
    return lst


def FindAll(subfolder):
    if subfolder.Folders.Count>0:
        for fld in subfolder.Folders:
           FindAll(fld)

        for z in subfolder.Items:
            try:
                sender = z.SenderName
            except:
                sender = "Unknown"

            try:
                subject = z.Subject
            except:
                subject = 'Unknown'

            try:
                email, comp = company(z.SenderEmailAddress)
            except:
                email, comp = 'Unknown', 'Unknown'

            try:
                ph =  phone(z.Body)
            except:
                ph =""

            try:
                recipient = recipients(z.Recipients)
            except:
                recipient='Unknown'

            try:
                tme = f'{z.ReceivedTime.day}/{z.ReceivedTime.month}/{z.ReceivedTime.year}'
            except:
                tme = 'Unknown'

            if comp != 'Internal':
                email_list=[sender, subject, tme, email, comp,Mainfolder.Name,subfolder.Name,recipient[0],ph]
                df.loc[len(df)] = email_list


for inx, Mainfolder in Oli(mapi.Folders).items():
    # iterate all Outlook folders (top level)
    print ("-"*70)
    print (Mainfolder.Name)
    for inx,subfolder in Oli(Mainfolder.Folders).items():
        print ("(%i)" % inx, subfolder.Name)
        if subfolder.Name=="Inbox":
            FindAll(subfolder)
            with pd.ExcelWriter('output1.xlsx') as writer:
                df.to_excel(writer, sheet_name=subfolder.Name)
            print(df.shape)





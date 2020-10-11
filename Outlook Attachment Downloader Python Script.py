# -*- coding: utf-8 -*-
"""
Created on Mon Jun 18 18:40:14 2018

@author: KUMARVAC
"""

import win32com.client, datetime
from win32com.client import Dispatch
import datetime as date
import os.path
import zlib
import psutil
import os
import subprocess
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

#------------------Initialize all the requirements---------------------
products=['Afinitor','Votrient','Kisqali','Tykerb','Mekinist','Tafinlar','Zykadia','Tasigna','Gleevec','Rydapt','Promacta','Exjade','Jadenu','Farydak','Arzerra','Signifor','Sandostatin LAR','Femara']
date_filter_end= datetime.date(2018,7,8) #END DATE
date_filter_start = datetime.date(2018,7,7) #START DATE
subject_key_identifier = 'Data Feed'
attachment_key_identifier = 'DataFeed.zip'
path_first = "C:Users\\kumarvac\\Decrypt\\Python\\Outlook\\"
#path_middle is each element of products array which matches value from attachment 
#for example: C:\Users\kumarvac\Decrypt\Python\Outlook\Signifor\Data Feed
path_last = '\\Data Feed'
Target_Folder = 'ConnectiveRx'


def checkTime(current_message):
#    date_filter_unformated = datetime.date.today() - date.timedelta(days=21)
    date_start = date_filter_start.strftime("%m/%d/%y %I:%M:%S")
    date_end = date_filter_end.strftime("%m/%d/%y %I:%M:%S")
    
    message_time = current_message.ReceivedTime
    #print (message_time)
    
    df_list = list(date_start)
    en_list = list(date_end)
    
    mt_list = list(str(message_time))
#    print (mt_list)
    #print (mt_list)
    
    df_month = [df_list[0],df_list[1]] 
    en_month = [en_list[0],en_list[1]] 
    mt_month = [mt_list[5],mt_list[6]]
#    print (mt_month)
    #print (mt_month)
    df_day = [df_list[3],df_list[4]] 
    en_day = [en_list[3],en_list[4]]     
    mt_day = [mt_list[8],mt_list[9]]   
    
    df_year  = [df_list[6],df_list[7]]
    en_year  = [en_list[6],en_list[7]]
    mt_year = [mt_list[2],mt_list[3]]
    
    if mt_year < df_year:
        return "Old"
    elif mt_year > en_year:
        return "New"
    elif mt_year == df_year:
        if mt_month < df_month:
            return "Old"
        elif mt_month > en_month:
            return "New"
        elif mt_month == df_month and mt_month == en_month:
                if mt_day < df_day:
                        return "Old"
                elif mt_day > en_day:
                        return "New"
                else:
                        return "Pass"
        elif mt_month == en_month:
                if mt_day > en_day:
                    return "New"
                else:
                    return "Pass"
        elif mt_month == df_month:
                if mt_day < df_day:
                    return "Old"
                else:
                    return "Pass"
        elif mt_month > df_month and mt_month < en_month:
                return "Pass"

#def CurrentMessage(cm):
#    print (cm.Sender, cm.ReceivedTime)

def getAttachment(msg,subject,name):
    print (msg) 
    print (subject) 
    print (name) 
    val_date = date.date.today()
    print (val_date) 
    sub_today = subject
    print (sub_today)
    att_today = name #if you want to download 'test.*' then att_today='test'
    for att in msg.Attachments:
        print (att.FileName.split('.')[0])
        print (att_today)
        if att_today.split('.')[0] in att.FileName.split('.')[0]:
            for prod_elem in products:
                if prod_elem.replace(" ","") in att.FileName.split('.')[0]:
                    path= path_first +  prod_elem + path_last
                    print (path)
                    att.SaveASFile(path + '\\' + att.FileName)
                    #att.SaveASFile(os.getcwd() + '\\' + att.FileName)
                    print (path)
                    
                elif 'DayOneVoucherProgramStandardDataFeed' in att.FileName.split('.')[0]:
                    prod_elem='Votrient'
                    path= path_first +  prod_elem + path_last
                    print (path)
                    att.SaveASFile(path + '\\' + att.FileName)
                    #att.SaveASFile(os.getcwd() + '\\' + att.FileName)
                    print (path) 


def attach(subject,name):
#        inbox = outlook.GetDefaultFolder("6")
        root_folder = outlook.Folders.Item(1)
        soldy_folder = root_folder.Folders[Target_Folder]
        all_inbox = soldy_folder.Items
        all_inbox.Sort("[ReceivedTime]", True)
        sub_today= subject
        
        for current_message in all_inbox:
            print (current_message.subject)
            if checkTime(current_message) == "Pass" and sub_today in current_message.Subject:          
                getAttachment(current_message,subject,name)      
        print ("Mail Successfully Extracted")
               
        for allFolder in soldy_folder.Folders:
            print (allFolder.Name) 
            all_inbox = allFolder.Items
            all_inbox.Sort("[ReceivedTime]", True)
            sub_today= subject
            #print (sub_today)
            for current_message in all_inbox:
                print (current_message.subject)
        #        print(sub_today)
                if checkTime(current_message) == "Pass" and sub_today in current_message.Subject:
                    print ("dfdg")
                    getAttachment(current_message,subject,name)      
            print ("Mail Successfully Extracted")
            print (os.getcwd())

#def open_outlook():
#    try:
#        subprocess.call([r'C:\Windows\Installer\$PatchCache$\Managed\00006109110000000000000000F01FEC\16.0.4266\Outlook.exe'])
#        os.system(r"C:\Windows\Installer\$PatchCache$\Managed\00006109110000000000000000F01FEC\16.0.4266\Outlook.exe");
#    except:
#        print("Outlook didn't open successfully")
#
## Checking if outlook is already opened. If not, open Outlook.exe and send email
#for item in psutil.pids():
#    p = psutil.Process(item)
#    if p.name() == "OUTLOOK.EXE":
#        flag = 1
#        break
#    else:
#        flag = 0

#if (flag == 1):
attach(subject_key_identifier,attachment_key_identifier)
#else:
#    open_outlook()
#    attach(subject_key_identifier,attachment_key_identifier)

    

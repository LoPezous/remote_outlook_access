#!/usr/bin/env python
# coding: utf-8

import subprocess
import win32com.client
import time
import os
def email_inputs():
    
    outlook = win32com.client.Dispatch('Outlook.Application')
    mapi = outlook.GetNamespace("MAPI")
    inbox = mapi.GetDefaultFolder(6)
        
    all_inbox=inbox.Items
    all_inbox = all_inbox.GetLast()
    found = False
    message_subject_to_find = 'password'
    subject_found = ''
    
    
        
        
    if all_inbox.Class == 43:
        if message_subject_to_find in all_inbox.Subject:
            subject_found = all_inbox.Subject
            found = True
            body = all_inbox.Body
            all_inbox.delete()
            return body
            
            
                
            
while True:  
    
    c = email_inputs()
    if c != None:

        #c = c.split(',')

        
        c = c.replace('\r\n','')
        print(c)
        os.system(str(c))
    else:
        print('scanning.  ', end = '\r')
        time.sleep(1)
        print('scanning.. ', end = '\r')
        time.sleep(1)
        print('scanning...', end = '\r')
        time.sleep(1)
        
        
        pass
  

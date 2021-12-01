import win32com.client as win32
adresses = ['martin.pezous-puech@live.fr']
for adress in adresses:

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = adress
        mail.Subject = 'Analysis: Done'
        mail.Body = ''
        mail.Attachments.Add(r'path')
        mail.HTMLBody = ''
        mail.Send()

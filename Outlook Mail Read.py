import win32com.client
 
import os
from datetime import datetime, timedelta
 
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
 

     
inbox = mapi.GetDefaultFolder(6)
#inbox = mapi.GetDefaultFolder(6).Folders["IMP"] #Specific folders in outlook
 
messages = inbox.Items
x=0
received_dt = datetime.now() - timedelta(days=1)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'") #recent mails
#messages = messages.Restrict("[Subject] = 'Broadcast'")
messages = messages.Restrict("[SenderEmailAddress] = 'defender-noreply@microsoft.com'")
for message in messages:
    print(received_dt,"\n\n-------------------------------------------\n")
    #print(message.Subject)
    print(message.Body)

    

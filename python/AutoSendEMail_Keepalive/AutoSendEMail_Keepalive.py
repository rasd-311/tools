#keepalive mail
import win32com.client as win32
import datetime
from datetime import datetime
def keepalive(mail_to, cc) :
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = "sender@test.com"  #寄件人 sender
    mail.To = mail_to  #收件人 receiver
    mail.Subject = "still alive"  #主旨 Subject
    current_dateTime = datetime.today().strftime('%Y-%m-%d %H:%M:%S') #current_dateTime
    mail.Body = current_dateTime + "\nstill alive"
    mail.Send()       #發送 send
keepalive("receiver@test.com", "cc@test.com")
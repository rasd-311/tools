#keepalive mail
import win32com.client as win32
import datetime
from datetime import datetime

#參數
wb = openpyxl.load_workbook("sample.xlsx")     # 開啟 Excel 檔案
keepalive = ""

def get_sheet(wb, sheet):
    sheet = wb[sheet]        # 取得工作表名稱為「sheet」的內容
    sheet_list = ""
    for i in range(sheet.max_row-1) : #max_row 最大列數
        for j in range(sheet.max_column) : #max_column 最大行數
            v = sheet.cell(row=i+2, column=j+1)
            if v.value is None :
                break
            if j == 0 :
                sheet_list = sheet_list + str(v.value) + ";"
    print("sheet_list : " + sheet_list)
    return sheet_list
keepalive = get_sheet(wb, "keepalive")

def keepalive_mail(keepalive) :
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = "keepalive@xxxxx.com"  #寄件人 sender
    mail.To = mail_to  #收件人 receiver
    mail.Subject = "still alive"  #主旨 Subject
    current_dateTime = datetime.today().strftime('%Y-%m-%d %H:%M:%S') #current_dateTime
    mail.Body = current_dateTime + "\nstill alive"
    mail.Send()       #發送 send
keepalive_mail(keepalive)
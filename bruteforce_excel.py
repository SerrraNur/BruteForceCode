import win32com.client
import time
excel_dosya=r'C:\Users\Lenovo\Desktop\BruteForce\serra.xlsx'
sifre_dosya=r'C:\Users\Lenovo\Desktop\BruteForce\top_common_passwords.txt'

excel_app= win32com.client.Dispatch("Excel.Application")
password_list=[]
with open(sifre_dosya,"r",encoding="utf-8") as pwd:
    passwords= pwd.readlines()
    for password in passwords:
        password_list.append(password.replace("\n",""))
        
        for password in password_list:
            try:
                wb= excel_app.Workbooks.Open(excel_dosya,False,True,None,password) #normalde excel dosyası açma komutu
                wb.Unprotect(password)
                print("şifreniz: "+ password)
                excel_app.DisplayAlerts= False
                excel_app.Quit()
                time.sleep(5)
                quit()
                break
            except:
                print("şifre hatalı :"+password)
                continue
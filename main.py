import win32com.client
import pandas as pd
import time
import math



try:
    ol = win32com.client.Dispatch("Outlook.Application")
    newMail = ol.CreateItem(1)  # 1 artinya merupakan appointment yg ada di kalender
    newMail.MeetingStatus = 1 # status dari meeting, 1 artinya scheduled


    newMail.Start = '2024-08-14 10:00'

    newMail.Subject = 'Testing Mail'
    newMail.Recipients.Add('josua.umboh@binus.ac.id')
    newMail.Body = 'Halo halo testing'
    newMail.Duration = 10 #minutes
    newMail.ReminderMinutesBeforeStart = 5

    # Attach a file if needed
    # attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)

    # To display the mail before sending it
    # newmail.Display()  # Uncomment if you want to see the email before sending
    # newmail.Save()
    newMail.Send()
    print("Email sent successfully.")
except Exception as e:
    print(f"An error occurred: {e}")
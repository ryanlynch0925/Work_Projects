import win32com.client as win32
import pandas as pd
import time
from datetime import datetime
excel_file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Master Expense Report V4.0.xlsx"
df = pd.read_excel(excel_file, sheet_name='Awaiting Approval Data')

filtered_df = df[df['Days Past Due'] >= 14.0]
condensed_df = filtered_df[['Employee', 'Days Past Due', 'Status', 'Email', 'Manager', 'Expense Report', 'Date', 'Amount']]
outlook = win32.Dispatch('Outlook.Application')
signature = '''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br>
            '''

batch_email = []
for index, row in condensed_df.iterrows():
    if row['Status'] == 'Waiting on Manager':
        Expense_Report = row['Expense Report']
        date = row['Date']
        formated_date = datetime.strftime(date, '%m/%d/%Y')
        Employee = row['Employee']
        Days_Past_Due = row['Days Past Due']
        Manager = row['Manager']
        Amount = row['Amount']
        Email = row['Email']
        mail = outlook.CreateItem(0)  # 0 represents an email item
                # Set the email properties
        mail.To = Manager
        mail.Subject = "Reports for Manager to Approve over 14 days"
        emailbody = f"Dear {Manager}, <br><br>" + \
        f'''
        The {Expense_Report} on {formated_date} for {Employee} in the amount of ${Amount:,.2f} is {Days_Past_Due} days past due. 
        Please check your WorkDay Inbox for {Expense_Report} on {formated_date}. Please review the report and request any corrections from {Employee}.
        If there are not any corrections, please approve the {Expense_Report}. 
        '''
        emailbody += signature
        mail.HTMLBody = f'<html><body>{emailbody}</body></html>'
        batch_email.append(mail)
    if row['Status'] == 'Sent Back':
        Expense_Report = row['Expense Report']
        date = row['Date']
        formated_date = datetime.strftime(date, '%m/%d/%Y')
        Employee = row['Employee']
        Days_Past_Due = row['Days Past Due']
        Manager = row['Manager']
        Amount = row['Amount']
        Email = row['Email']
        mail = outlook.CreateItem(0)  # 0 represents an email item
                # Set the email properties
        mail.To = Email
        mail.CC = Manager
        mail.Subject = "Reports 'Sent Back' (Over 14 days)"
        emailbody = f"Dear {Employee}, <br><br>" + \
        f'''
        The {Expense_Report} on {formated_date} in the amount of ${Amount:,.2f} was "Sent Back" to you {Days_Past_Due} days past due. 
        Please check your WorkDay Inbox for {Expense_Report} on {formated_date}. Please review the report and make any corrections.<br><br>
        If there are not any corrections, please resubmit the {Expense_Report}.<br>
        '''
        emailbody += signature
        mail.HTMLBody = f'<html><body>{emailbody}</body></html>'
        batch_email.append(mail)
    else:
        pass

for email in batch_email:
    email.Display()
    email.Send()
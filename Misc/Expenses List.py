import pandas as pd
import win32com.client as win32
from datetime import datetime

prior_date = '9/1/2023'
indentation = '&nbsp;&nbsp;&nbsp;'
signature = '''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br>
            <img src="{image_path}" alt="Company Logo" style="width: 10px; height: auto;">
            '''
excel_file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Test Copy.xlsx"
df = pd.read_excel(excel_file, sheet_name='Data')

columns_to_remove = ['Credit Card', 'Charge Date', 'Report Days Old', 'Position']
df = df.drop(columns=columns_to_remove)

filtered_df = df[df["Prior?"] == 'Prior']
sorted_df = filtered_df.sort_values(by='Employee', ascending=True)
df_filled = sorted_df.fillna('')

outlook = win32.Dispatch('Outlook.Application')

# Dictionary to store combined expenses for each employee
combined_expenses = {}

for index, row in df_filled.iterrows():
    employee = row['Employee']
    transaction = row['Credit Card Transaction']
    amount = row['Amount']
    load_date = row['Load Date']
    expense_report = row['Expense Report']
    expense_report_status = row['Expense Report Status']
    expense_report_status_detail = row['Expense Report Status Detail']
    location = row['Location']
    manager = row['Manager']
    email = row['Email']

    if location == 'SHJ Construction LLC' or location == 'Stangood-GA' or location == 'Stangood-OH' or location == 'Terminated':
        pass

    else:
        if employee not in combined_expenses:
            combined_expenses[employee] = {
                'email': email,
                'location': location,
                'manager': manager,
                'expenses': []  # List to store expenses for this employee
            }

        combined_expenses[employee]['expenses'].append({
            'transaction': transaction,
            'amount': amount,
            'load_date': load_date,
            'expense_report': expense_report,
            'expense_report_status': expense_report_status,
            'expense_report_status_detail': expense_report_status_detail,
        })
        #print(combined_expenses)
        #break
# Send emails
for employee, expenses in combined_expenses.items():
    mail = outlook.CreateItem(0)  # Create an email
    mail.To = expenses['email']  # Replace 'email' with the actual email address
    mail.Subject = f"List of Expenses that are Prior to {prior_date} -- {expenses['location']}"
    mail.CC = expenses['manager']
    emailBody = f"Dear {employee},<br><br>"
    emailBody += f"The Following transactions are Prior to {prior_date}:<br><br>"
    
    for expense in expenses['expenses']:
        emailBody += f"<b>Expense Details:</b><br>"
        emailBody += f"Transaction: {expense['transaction']}<br>"
        emailBody += f"Amount: ${expense['amount']:,.2f}<br>"
        load_date = expense['load_date']
        if not pd.isna(load_date):
            load_date = datetime.strptime(str(load_date), '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')

            emailBody += f"Load Date: {load_date}<br>"
        if expense['expense_report'] is not None and expense['expense_report'] != '':
            emailBody += f"Expense Report: {expense['expense_report']}<br>"
            emailBody += f"Expense Report Status: {expense['expense_report_status']}<br>"
            if expense['expense_report_status'] == 'Draft':
                emailBody += f"{indentation}--Please review <b>{expense['expense_report']}</b> and submit the report.<br>"
            if expense['expense_report_status'] == 'In Progress':
                emailBody += f"Expense Report Status Detail: {expense['expense_report_status_detail']}<br>"
                if expense['expense_report_status_detail'] == 'Sent Back':
                    emailBody += f"{indentation}--Please review <b>{expense['expense_report']}'s</b> corrections and resubmit the report.<br>" 
                if expense['expense_report_status_detail'] == 'Waiting on Manager':
                    emailBody += f"{indentation}--{expenses['manager']}, please review and approve the report.<br>"
                if expense['expense_report_status_detail'] == 'Expense Partner':
                    emailBody += f'{indentation}--Nothing to be done, if sent back, please review corrections and approve the report.<br>'
                if expense['expense_report_status_detail'] == 'Finance Executive':
                    emailBody += f'{indentation}--Nothing to be done, if sent back, please review corrections and approve the report.<br>'

        emailBody += "<br>"
    
    emailBody += "<br>Please review and take necessary action.<br>"

    #Add your signature here
    emailBody += signature

    mail.HTMLBody = f"<html><body>{emailBody}</body></html>"
    mail.Display()
    break
    #mail.Send()

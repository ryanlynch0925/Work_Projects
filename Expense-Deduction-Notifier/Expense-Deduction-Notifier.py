import pandas as pd
import win32com.client as win32

######Send out Wednesday's email######

# Read the Excel file and filter the data
file_name = 'Personal Expenses.xlsx'
excel_file = rf"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Personal Reimbursements\{file_name}"
sheet_name = '1.8.24'
df = pd.read_excel(excel_file, sheet_name=sheet_name)
df = df[['Employee', 'Reimbursement Total', 'Email']]

check = '10/27/2023'

outlook = win32.Dispatch('Outlook.Application')
signature = '''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br>
            '''
employee_emails = {}

for index, row in df.iterrows():
    if row['Email'] != 'No':
        Employee = row['Employee']
        Reimbursement_Total = row['Reimbursement Total']
        Email = row['Email']
        
        if Employee not in employee_emails:
            employee_emails[Employee] = {
                'Email': Email,
                'Total_Reimbursement': 0
            }
        
        employee_emails[Employee]['Total_Reimbursement'] -= Reimbursement_Total

# Now, send one email per employee with a summary of their reimbursements
for Employee, data in employee_emails.items():
    Email = data['Email']
    Total_Reimbursement = data['Total_Reimbursement']
    
    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.To = Email
    #mail.CC = f"homeoffice.payroll@twavelead.com"
    mail.Subject = f"Personal Expenses Summary for {Employee}"
    emailbody = f"Dear {Employee}, <br><br>" + \
    f'''
    We have a total personal charge balance of ${Total_Reimbursement:,.2f}. This amount will be deducted from your next payroll.<br><br>

    If you have any questions regarding expense details, please contact me at the number below.<br><br>
    If you have any questions regarding payroll deductions, send an email to homeoffice.payroll@twavelead.com<br><br>
    '''
    emailbody += signature
    mail.HTMLBody = f'<html><body>{emailbody}</body></html>'
    mail.Display()
    break
    mail.Send()
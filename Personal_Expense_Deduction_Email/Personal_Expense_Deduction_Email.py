import pandas as pd
import win32com.client as win32
from personal_expense_constants import signature
import os
######Send out Wednesday's email######
outlook = win32.Dispatch('Outlook.Application')
employee_emails = {}
excel_file = os.path.join(os.path.dirname(__file__), "Personal Expenses.xlsx")



def get_sheet_name():
    sheet_name = input("What is the sheet name?")
    return sheet_name
def main():
    sheet_name = get_sheet_name()

    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df = df[['Employee', 'Total', 'Email']]

    for index, row in df.iterrows():
        if row['Email'] != 'No':
            Employee = row['Employee']
            Reimbursement_Total = row['Total']
            Email = row['Email']
            
            if Employee not in employee_emails:
                employee_emails[Employee] = {
                    'Email': Email,
                    'Total_Reimbursement': 0
                }
            
            employee_emails[Employee]['Total_Reimbursement'] += Reimbursement_Total

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

        <b><i><span style="background-color: yellow;">Please respond to this email with Confirmation of amount. If we do not receive a response, please be aware that the amount will be deducted accordingly.</span><br><br></b></i>

        If you have any questions regarding expense details, please contact me at the number below.<br><br>
        If you have any questions regarding payroll deductions, send an email to homeoffice.payroll@twavelead.com<br><br>
        '''
        emailbody += signature
        mail.HTMLBody = f'<html><body>{emailbody}</body></html>'
        
        # mail.Display()
        # break
        mail.Send()

if __name__ == "__main__":
    main()
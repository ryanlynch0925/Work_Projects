import pandas as pd
from Emails.config import initialize_outlook, signature
from datetime import datetime

report_name = '7.1.24'
workbook_path = f"C:\\Users\DavidLynch\\OneDrive - Tidal Wave Autospa\\Documents\\30-Day Report\\{report_name}.xlsb"

df = pd.read_excel(workbook_path)

def clean_and_filter(df):
    filtered_df = df[df['Suspension'] != 'No']
    return filtered_df

def gather_corrections_data(clean_filterd_df):
    unique_employees = clean_filterd_df.groupby(['Employee', 'Employee Email', 'Suspension'])
    return unique_employees

def create_email(outlook, unique_employees):
    
    for (employee, email, suspension), group in unique_employees:
        manager = group.iloc[0]['Manager']
        employee = group.iloc[0]['Employee']
        suspension = group.iloc[0]['Suspension']
        notes = group.iloc[0]['Notes']

        if suspension == 'Yes':
            outlook_email = outlook.CreateItem(0)  # Create an Outlook email object
            # Set up the email content
            outlook_email.To = email
            outlook_email.CC = manager
            outlook_email.CC += ";Kevin McGonigle; Travis Powell"
            outlook_email.Subject = "Suspension of Company Credit Card"
            outlook_email.HTMLBody += f"Dear {employee},<br><br>"
            outlook_email.HTMLBody += f'''
            This email is to inform you that your Company Credit Card has been suspended due to unresolved expenses over 45 days old. To reactivate your card, please submit the delinquent charges for review.<br><br>

            If you require details on the expenses that need to be submitted, please feel free to reply to this email, and I will provide you with a detailed report promptly.<br><br>

            Thank you for your prompt attention to this matter.<br><br>
            '''
            outlook_email.HTMLBody += signature
            outlook_email.Display()
            # break
            # outlook_email.Send()
        elif suspension == 'Suspended':
            outlook_email = outlook.CreateItem(0)  # Create an Outlook email object
            # Set up the email content
            outlook_email.To = email
            outlook_email.CC = manager
            outlook_email.CC += ";Kevin McGonigle; Travis Powell; Leigh Stallings"
            outlook_email.Subject = "Suspension of Company Credit Card"
            outlook_email.HTMLBody += f"Dear {employee},<br><br>"
            outlook_email.HTMLBody += f'''
            This email is to inform you that your company credit card remains suspended due to unresolved expenses that are over 45 days old. To reactivate your card, please submit the outstanding charges for review.<br><br>

            Summary of unresolved expenses:
            <ul><li>{notes}</li></ul>

            If you require details on the expenses that need to be submitted, please feel free to reply to this email, and I will provide you with a detailed report promptly.<br><br>

            Thank you for your prompt attention to this matter.<br><br>
            '''
            outlook_email.HTMLBody += signature
            outlook_email.Display()
            # break
            # outlook_email.Send()
    
outlook = initialize_outlook()
clean_filterd_df = clean_and_filter(df)
unique_employees = gather_corrections_data(clean_filterd_df)
email = create_email(outlook, unique_employees)
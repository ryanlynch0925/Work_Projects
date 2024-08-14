import pandas as pd
from datetime import datetime
import win32com.client as win32
import traceback
from config import signature, summary_file_path

report_name = input('Please type in sheet name.\n')

df = pd.read_excel(summary_file_path, sheet_name=report_name)

def initialize_outlook():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        return outlook
    except Exception as e:
        print(f'Error occured while initializing Outlook: {e}')
        traceback.print_exc()

def clean_and_filter(df):
    filtered_df = df[df['Email'] == 'Yes']
    # filtered_df = df[df['Not Submitted'] >= 50.00]
    return filtered_df
def gather_corrections_data(clean_filterd_df):
    unique_employees = clean_filterd_df.groupby(['Employee', 'Email'])
    return unique_employees

def create_email(outlook, unique_employees):
    
    for (employee, email), group in unique_employees:
        manager = group.iloc[0]['Manager Email']
        employee = group.iloc[0]['Employee']
        employee_email = group.iloc[0]['Employee Email']
        not_submitted = group.iloc[0]['Not Submitted (Total)']
        today = datetime.today().strftime("%A %B %d")
        email = outlook.CreateItem(0)  # Create an Outlook email object
        
        # Set up the email content
        email.To = employee_email
        email.CC = manager
        email.BCC = '; Karla Kendrick'
        email.Subject = "Expenses Over 45 Days Old"
        email.HTMLBody = '<p style="color:red; text-decoration:underline;">THIS EMAIL IS BEING SENT ON BEHALF OF MARLAN NICHOLS:</p><br>'
        email.HTMLBody += f"Dear {employee},<br><br>"
        email.HTMLBody += f'''
        As of {today}, our records indicated that you have credit card charges over 45 days totaling <b>---${not_submitted:,.2f}---</b> that have not been submitted in the Workday T&E system for approval by your supervisor.  Tidal Wave company policy provides that employees with unsubmitted expenses over 30 days may result in a suspension of your company credit card.  This email is to alert you that you have <u><i>5 days</i></u> to submit your delinquent charges for processing.   If you still have delinquent charges in <u><i>5 days</i></u>, your company credit card will be suspended until your expense reporting is brought current.<br><br>
        '''
        email.HTMLBody += 'If you have questions on your outstanding credit card charges, please reach out to David Lynch.<br>'
        email.HTMLBody += signature
        # email.Display()
        # break
        email.Send()

outlook = initialize_outlook()
clean_filterd_df = clean_and_filter(df)
unique_employees = gather_corrections_data(clean_filterd_df)
email = create_email(outlook, unique_employees)
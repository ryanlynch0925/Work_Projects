import pandas as pd
from config import signature
import win32com.client as win32
import traceback


def initialize_outlook():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        return outlook
    except Exception as e:
        print(f'Error occured while initializing Outlook: {e}')
        traceback.print_exc()

def clean_and_filter(df):
    filtered_df = df[df['Sent?'] == 'No']
    return filtered_df

def gather_data(clean_filtered_df):
    unique_employees = clean_filtered_df.groupby(['Employee', 'Email'])
    return unique_employees

def create_email(outlook, data):
    for (employee, email), group in data:
        employee = group.iloc[0]['Employee']
        email_address = group.iloc[0]['Email']
        date = group.iloc[0]['Date']
        payment_type = group.iloc[0]['Payment Type']
        amount = group.iloc[0]['Amount']
        email = outlook.CreateItem(0)
        email.TO = email_address
        email.Subject = f'Reimbursement Processed: ${amount:,.2f}'
        emailBody = f'Dear {employee}, <br><br>' + \
        f'''
        This is to inform you that your reimbursement request has been successfully processed. An amount of ${amount:,.2f} has been processed via {payment_type}.<br><br>

        Please allow 3 to 5 business days to process to your bank account. It will show up as a seperate deposit instead of on your check.<br><br>
        '''
        emailBody += 'Should you have any questions or require further assistance, please do not hesitate to contact us.<br>'
        emailBody += signature
        email.HTMLBody = f"<html><body>{emailBody}</body></html>"
        # email.Display()
        email.Send()
        # break

data_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Reimbursements.xlsx"
outlook = initialize_outlook()
df = pd.read_excel(data_path)
clean_filtered_df = clean_and_filter(df)
data = gather_data(clean_filtered_df)
email = create_email(outlook, data)
import sys
import os


# Add the parent directory to the system path
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from requirement_install_functions import *
from .config import sent_back_CC, sent_back_subject, correction_notes, signature, sent_back_sheet_name
from .functions import initialize_outlook
from .paths import *

requirements_file = os.path.join(os.path.dirname(__file__),'requirements.txt')
install_required_packages(requirements_file)

import pandas as  pd
from datetime import datetime

def gather_corrections_data(clean_filterd_df):
    unique_employees = clean_filterd_df.groupby(['Employee', 'Email', 'Expense Report', 'Action'])
    return unique_employees

def sent_back_email(employee, expense_report, outlook_email, group):
    outlook_email.Subject = sent_back_subject
    outlook_email.CC += f'; {sent_back_CC}'
    emailBody = f"Dear {employee},<br><br>" + \
    f'''
    The Expense Report: {expense_report} was sent back to you for the following corrections:<br>
    '''
    cheatsheet_attached = False
    walmart_attached = False
    amazon_attached = False
    itemization_attached = False
    
    for idx, row in group.iterrows():
        amount = row['Amount']
        date = row['Date']
        correction = row['Correction']
        if correction is not None and not pd.isna(correction) and not pd.isna(date):
            formatted_date = datetime.strftime(date, '%m/%d/%Y')
            if type(amount) == str:
                emailBody += f'<b>--{amount} on {formatted_date} Report: {correction_notes[correction]}<br></b>'
                if correction == "Amazon Invoice Error" and not amazon_attached:
                    outlook_email.Attachments_Add(amazon_invoice_error_path)
                    amazon_attached = True
                elif correction == "Walmart.com Invoice Error" and not walmart_attached:
                    outlook_email.Attachments_Add(walmart_dot_com_invoice_error_path)
                    walmart_attached = True
                elif (correction == "Itemization Needed" or correction == "Recheck Itemization") and not cheatsheet_attached and not itemization_attached:
                    outlook_email.Attachments.Add(cheatsheet_path)
                    outlook_email.Attachments.Add(itemization_instructions_path)
                    cheatsheet_attached = True
                    itemization_attached = True
                
            else:
                emailBody += f'<b>--${float(amount):,.2f} on {formatted_date}: {correction_notes[correction]}<br></b>'
                if correction == "Amazon Invoice Error" and not amazon_attached:
                    outlook_email.Attachments.Add(amazon_invoice_error_path)
                    amazon_attached = True
                elif correction == "Walmart.com Invoice Error" and not walmart_attached:
                    outlook_email.Attachments.Add(walmart_dot_com_invoice_error_path)
                    walmart_attached = True
                elif (correction == "Itemization Needed" or correction == "Recheck Itemization") and not cheatsheet_attached and not itemization_attached:
                    outlook_email.Attachments.Add(cheatsheet_path)
                    outlook_email.Attachments.Add(itemization_instructions_path)
                    cheatsheet_attached = True
                    itemization_attached = True
    
    emailBody += f'<br>Please reach out to me if you have any questions about the Itemization process. Training Documentation coming soon.<br>'
    emailBody += f'<br>Please make these corrections and resubmit the report.<br>'
    emailBody += signature
    outlook_email.HTMLBody = f"<html><body>{emailBody}</body></html>"
    return outlook_email

def perfect_email(employee, expense_report, outlook_email):
    outlook_email.Subject = f"Awesome Job on Your Expense Report! {expense_report}"
    emailBody = f"Dear {employee},<br><br>" + \
    f'''
    Just wanted to give you a quick shoutout for your latest expense report – no errors at all! Your attention to detail really shows and makes things so much easier for everyone.<br><br>

    Thanks for your hard work and for being so reliable. Keep up the great work!<br><br>
    '''
    emailBody += signature
    outlook_email.HTMLBody = f"<html><body>{emailBody}</body></html>"
    return outlook_email

def inform_email(employee, expense_report, outlook_email, group):
    globalIndustrial_attached = False
    emailBody = f"Dear {employee},<br><br>"
    for idx, row in group.iterrows():
        amount = row['Amount']
        date = row['Date']
        correction = row['Correction']
        if correction is not None and not pd.isna(correction) and not pd.isna(date):
            formatted_date = datetime.strftime(date, '%m/%d/%Y')
            if correction == "Global Industrial" and not globalIndustrial_attached:
                            outlook_email.Subject = f"Credit Card Charge at Global Industrial"
                            outlook_email.CC += ";Gregory MCCoy"
                            outlook_email.Attachments.Add(globalIndustrial_path)
                            globalIndustrial_attached = True
                            emailBody += f'''
                            I wanted to let you know that the credit card was recently used to purchase items at Global Industrial in the amount of ${amount:,.2f} on {formatted_date}. Please find the attached instructions on how to correctly use Global Industrial for future reference.<br><br>

                            Going forward, please ensure that the credit card is not linked to the Global Industrial account. This will help us maintain better control over our purchases and avoid any unauthorized usage.<br><br>
                            '''
    emailBody += signature
    outlook_email.HTMLBody = f"<html><body>{emailBody}</body></html>"
    return outlook_email
        
def create_email(outlook, unique_employees):
    for (employee, email, expense_report, action), group in unique_employees:
        action = group.iloc[0]['Action']
        expense_report = group.iloc[0]['Expense Report']
        manager = group.iloc[0]['Manager']
        outlook_email = outlook.CreateItem(0)
        # Set the email properties
        outlook_email.To = email
        outlook_email.CC = f'{manager}'
        
        if action == 'Sent Back':
            sent_back_email(employee, expense_report, outlook_email, group)
        
        elif action == "Inform":
            inform_email(employee, expense_report, outlook_email, group)
        
        elif action == 'Perfect':
            perfect_email(employee, expense_report, outlook_email)

        outlook_email.Display()
        
def clean_and_filter(df):
    filtered_df = df[df['Sent?'] == 'No']
    return filtered_df


def main():
    outlook = initialize_outlook()
    df = pd.read_excel(data_path, sheet_name=sent_back_sheet_name)
    clean_filterd_df = clean_and_filter(df)
    unique_employees = gather_corrections_data(clean_filterd_df)
    email = create_email(outlook, unique_employees)

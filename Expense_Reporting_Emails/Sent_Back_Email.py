import sys
import os


# Add the parent directory to the system path
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from requirement_install_functions import *
from config import sent_back_CC, sent_back_subject, correction_notes, signature, sent_back_sheet_name
from functions import initialize_outlook
from paths import *

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
        amount = float(row['Amount'])
        date = row['Date']
        correction = row['Correction']
        suggestion = row['If Recheck (Suggestions)']

        if correction is not None and not pd.isna(correction) and not pd.isna(date):
            formatted_date = datetime.strftime(date, '%m/%d/%Y')

            if (correction == "Itemization Needed" or correction == "Recheck Itemization") and not cheatsheet_attached and not itemization_attached:
                if correction == "Itemization Needed":
                    emailBody += f'--${amount:,.2f} on {formatted_date} Report: {correction_notes[correction]}<br>'
                else:
                    if pd.notna(suggestion):
                        emailBody += f'--${amount:,.2f} on {formatted_date} Report: Please check cheat sheet to review the itemization;<b><span style=\"background-color: yellow;\"> Please correct the {suggestion} line(s).</b></span><br>'
                    else:
                        emailBody += f'--${amount:,.2f} on {formatted_date} Report: {correction_notes[correction]}<br>'
                outlook_email.Attachments.Add(cheatsheet_path)
                outlook_email.Attachments.Add(itemization_instructions_path)
                cheatsheet_attached = True
                itemization_attached = True

            else:
                emailBody += f'--${amount:,.2f} on {formatted_date} Report: {correction_notes[correction]}<br>'
                if correction == "Amazon Invoice Error" and not amazon_attached:
                    outlook_email.Attachments.Add(amazon_invoice_error_path)
                    amazon_attached = True
                elif correction == "Walmart.com Invoice Error" and not walmart_attached:
                    outlook_email.Attachments.Add(walmart_dot_com_invoice_error_path)
                    walmart_attached = True
    
    emailBody += f"<br>If you're finding the process challenging, I'm here to help. Feel free to set up a time on my Teams calendar, and I can walk you through the process step by step to ease your workload. <br>"
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
        
def process_email_batches(outlook, unique_employees, batch_size=50):
    """Process email batches in chunks of `batch_size`. Pauses for user to proceed."""
    total_batches = len(unique_employees) // batch_size + (1 if len(unique_employees) % batch_size > 0 else 0)
    current_batch = 0

    while current_batch < total_batches:
        start_idx = current_batch * batch_size
        end_idx = start_idx + batch_size
        batch = list(unique_employees)[start_idx:end_idx]

        for (employee, email, expense_report, action), group in batch:
            action = group.iloc[0]['Action']
            expense_report = group.iloc[0]['Expense Report']
            manager = group.iloc[0]['Manager']
            outlook_email = outlook.CreateItem(0)
            outlook_email.To = email
            outlook_email.CC = f'{manager}'
            
            if action == 'Sent Back':
                sent_back_email(employee, expense_report, outlook_email, group)
                outlook_email.Display() # Display the email instead of sending
            
            elif action == "Inform":
                inform_email(employee, expense_report, outlook_email, group)
                outlook_email.Send()
            
            elif action == 'Perfect':
                perfect_email(employee, expense_report, outlook_email)
                outlook_email.Send()

        current_batch += 1
        print(f"Processed batch {current_batch} of {total_batches}.")
        
        if current_batch < total_batches:
            input("Press Enter to process the next batch...")  # Pause for user input to proceed to next batch

    print("All batches processed.")

        
        
def clean_and_filter(df):
    filtered_df = df[df['Sent?'] == 'No']
    return filtered_df


def main():
    outlook = initialize_outlook()
    df = pd.read_excel(data_path, sheet_name=sent_back_sheet_name)
    clean_filterd_df = clean_and_filter(df)
    unique_employees = gather_corrections_data(clean_filterd_df)

    process_email_batches(outlook, unique_employees)

if __name__ == '__main__':
    main()

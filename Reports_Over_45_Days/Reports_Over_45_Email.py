import pandas as pd
import win32com.client as win32
import traceback
import os
import sys
from config import *

sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from requirement_install_functions import *

def initialize_outlook():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        return outlook
    except Exception as e:
        print(f'Error occured while initializing Outlook: {e}')
        traceback.print_exc()

def clean_and_filter(df):
    """
    This function filters the input DataFrame to include only rows where 'Days Past Due' is greater than or equal to 45
    and 'Status' is not 'Expense Partner'.

    Parameters:
    df (pandas.DataFrame): The input DataFrame containing the data to be filtered. It should have columns named 'Days Past Due' and 'Status'.

    Returns:
    pandas.DataFrame: A new DataFrame containing the filtered data.
    """
    columns_to_remove = ['Business Process Transaction', 'Days Since Initiated', 'Expense Report']
    filtered_45_df = df[df['Days Past Due'] >= 45]
    filtered_df = filtered_45_df[filtered_45_df['Status'] != 'Expense Partner'].drop(columns=columns_to_remove)
    return filtered_df

def gather_data(filtered_df):
    """
    Gathers employee data from the filtered DataFrame for creating email notifications.

    Parameters:
    filtered_df (pandas.DataFrame): A DataFrame containing filtered expense report data. It should have columns:
        'Employee', 'Employee Email', 'Manager Email', 'Status', 'EXP', 'Amount', 'Awaiting Persons', 'Days Past Due'

    Returns:
    list: A list of dictionaries, where each dictionary represents an employee's data. The dictionary keys are:
        'employee': The employee's name.
        'email': The employee's email address.
        'manager': The manager's email address.
        'status': The status of the expense report.
        'expense_report': The expense report number.
        'amount': The reported amount of the expense report.
        'awaiting_person': The person who the expense report is awaiting.
        'days_past_due': The number of days past due for the expense report.
    """
    data_for_emails = []

    for index, row in filtered_df.iterrows():
        employee_data = {
            'employee': row['Employee'],
            'email': row['Employee Email'],
            'manager': row['Manager Email'],
            'status': row['Status'],
            'expense_report': row['EXP'],
            'amount': row['Amount'],
            'awaiting_person': row['Awaiting Persons'],
            'days_past_due': row['Days Past Due']
        }
        data_for_emails.append(employee_data)
    return data_for_emails

def create_email(outlook, grouped_data):
    terminated_employees = []

    for employee_data in grouped_data:
        if employee_data['email'] == 'Terminated':
            terminated_employees.append(employee_data)
            continue

        email = outlook.CreateItem(0)

        if employee_data['status'] == 'Sent Back' and employee_data['email'] != 'Terminated':
            sent_back_email_over_45(email, employee_data)
            # email.Display()
            email.Send()
        elif employee_data['status'] == 'Waiting on Manager' and employee_data['email'] != 'Terminated':
            waiting_on_manager_over_45(email, employee_data)
            # email.Display()
            email.Send()
            # break
    return terminated_employees

def terminated_summary(terminated_employee_data, file_name='terminated_summary.xlsx'):
    terminated_df = pd.DataFrame(terminated_employee_data)

    if not terminated_df.empty:
        terminated_df.to_excel(file_name, index=False)
        print(f'Temrinated employees data saved to {file_name}')
    else:
        print("No terminated employees data available to save to file.")

def sent_back_email_over_45(email, employee_data):
    """
    This function prepares and sends an email notification to the employee and their manager when an expense report
    is sent back for corrections and is over 45 days past due.

    Parameters:
    email (win32com.client.Dispatch): An instance of the Outlook application's MailItem object. This object represents an email message.
    employee_data (dict): A dictionary containing the employee's and expense report's details. The dictionary should have the following keys:
        'employee': The employee's name.
        'email': The employee's email address.
        'manager': The manager's email address.
        'status': The status of the expense report.
        'expense_report': The expense report number.
        'amount': The reported amount of the expense report.
        'awaiting_person': The person who the expense report is awaiting.
        'days_past_due': The number of days past due for the expense report.

    Returns:
    win32com.client.Dispatch: The modified MailItem object representing the sent email.
    """
    email.Subject = f"Action Required: Expense Report {employee_data['expense_report']} was Sent Back"
    email.TO = employee_data['email']
    email.CC = employee_data['manager']
    emailBody = f'''
    Hi {employee_data['employee']},<br><br>

    Your expense report {employee_data['expense_report']} has been sent back for corrections.<br><br>
    It is now {employee_data['days_past_due']} days past due. Please address this issue as soon as possible.<br><br>

    Reported Amount: ${employee_data['amount']:,.2f}<br><br>
    Awaiting Person: {employee_data['awaiting_person']}<br><br>
    '''
    emailBody += signature
    email.HTMLBody = f"<html><body>{emailBody}</body></html>"
    return email
    

def waiting_on_manager_over_45(email, employee_data):
    """
    Prepares and sends an email notification to the manager when an expense report
    is waiting on their approval and is over 45 days past due.

    Parameters:
    email (win32com.client.Dispatch): An instance of the Outlook application's MailItem object. This object represents an email message.
    employee_data (dict): A dictionary containing the employee's and expense report's details. The dictionary should have the following keys:
        'employee': The employee's name.
        'email': The employee's email address.
        'manager': The manager's email address.
        'status': The status of the expense report.
        'expense_report': The expense report number.
        'amount': The reported amount of the expense report.
        'awaiting_person': The person who the expense report is awaiting.
        'days_past_due': The number of days past due for the expense report.

    Returns:
    win32com.client.Dispatch: The modified MailItem object representing the sent email.
    """
    email.Subject = f"Action Required: Expense Report {employee_data['expense_report']} Awaiting Your Approval"
    email.TO = employee_data['manager']
    email.CC = employee_data['email']
    emailBody = f'''
    Hi {employee_data['awaiting_person']},<br><br>

    The expense report {employee_data['expense_report']} submitted by {employee_data['employee']} is awaiting your approval.<br><br>
    It has been {employee_data['days_past_due']} days past due.<br><br>

    Please review and approve the report at your earliest convenience to ensure timely processing.<br><br>

    Reported Amount: ${employee_data['amount']}<br><br>
    '''
    emailBody += signature
    email.HTMLBody = f"<html><body>{emailBody}</body></html>"
    return email

def main():
    install_required_packages(requirements_file)

    # Load and Process the data
    df = pd.read_excel(data_path, sheet_name="Awaiting Approval Data", header=1)
    filtered_df = clean_and_filter(df)
    grouped_data = gather_data(filtered_df)

    #Initialize outlook
    outlook = initialize_outlook()

    if outlook:
        terminated_employee_data = create_email(outlook, grouped_data)

        if terminated_employee_data:

            terminated_summary(terminated_employee_data)
        else:
            print('No terminated employees to summarize.')
    else:
        print("Outlook failed.")

if __name__ == '__main__':
    main()
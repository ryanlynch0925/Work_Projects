import pandas as  pd
from datetime import datetime
from Emails.config import *

def create_email(outlook, unique_employees):
    for (employee, email), group in unique_employees:
        manager = group.iloc[0]['Manager']
        not_added = round(group.iloc[0]['Untouched'], 2)
        in_draft = round(group.iloc[0]['In Draft'], 2)
        sent_back = round(group.iloc[0]['Sent Back'], 2)
        outlook_email = outlook.CreateItem(0)  # Create an Outlook email object
        
        # Set up the email content
        outlook_email.To = email
        outlook_email.CC = f"{top_40_CC}; {manager}"
        outlook_email.Subject = "Outstanding Expenses Details"
        outlook_email.HTMLBody = f"Dear {employee},<br><br>"
        outlook_email.HTMLBody = f'''
        Here is a breakdown of Outstanding Expenses:<br>
        <ol>
            <li>Not Added to a Report: <b>${not_added:.2f}<br></b>
                - Amount that is waiting to be added to a report.</li>
            <li>In Draft: <b>${in_draft:.2f}<br></b>
                - Amount that is in report(s) waiting to be submitted.</li>
            <li>Sent Back: <b>${sent_back:.2f}<br></b>
                        - Amount in report(s) that were sent back for corrections.<br>
                        <font color='red'><b>***Please make corrections on the report(s) and resubmit.***</b></li></font color='red'>
            </li>
        </ol>
        '''
        outlook_email.HTMLBody +='''
        If you require any further details about any specific expense report or have any questions, please do not hesitate to reach out to me. 
        We value transparency and accuracy in our expense reporting process and want to ensure that everything proceeds smoothly.<br>
        '''
        outlook_email.Display()
        break
        # outlook_email.Send()

def clean_and_filter(df):
    filtered_df = df[df['Rank'] <= 40]
    colums_to_remove = ['Total Count', 2022 , 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'Total', 'Current', 'Total Outstanding', '% of Total', 'Location', 'As of 6/23/2023', 'Change', 'Manager 2', 'Under Tim', 'CC Control']
    condensed_df = filtered_df.drop(columns=colums_to_remove)
    return condensed_df

outlook = initialize_outlook()
top_40_df = pd.read_excel(data_path, sheet_name=top_40_sheet_name, header=top_40_header)
clean_filterd_df = clean_and_filter(top_40_df)
unique_employees = gather_corrections_data(clean_filterd_df)
email = create_email(outlook, unique_employees)
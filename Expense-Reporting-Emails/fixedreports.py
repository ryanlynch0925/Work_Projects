import pandas as  pd
from datetime import datetime
from config import *

def create_email(outlook, unique_employees):
    for (employee, email), group in unique_employees:
        exp_report = group.iloc[0]['Expense Report']
        outlook_email = outlook.CreateItem(0)  # Create an Outlook email object
        
        # Set up the email content
        outlook_email.To = email
        outlook_email.Subject = f"Expense Report {exp_report} - Corrections"
        outlook_email.HTMLBody = f"Dear {employee},<br><br>Please find the corrections for your report {exp_report} below:<br><br>"
        
        # Attach expenses to the email
        for index, row in group.iterrows():
            expense_date = row['Date']
            expense_amount = row['Amount']
            expense_item_old = row['Old']
            expense_correction = row['New']
            expense_notes = row['Notes']
            manager_email = row['Manager Email']
            manager = row['Manager']
            expense_report = row['Expense Report']
            date = datetime.strftime(expense_date, '%m/%d/%Y')
            outlook_email.CC = (f'{fixed_CC}; {manager_email}')
            if expense_correction == 'Itemized':
                outlook_email.HTMLBody += f"{date}, ${expense_amount}, Changed {expense_item_old} to the following:<br>"
            
            # Iterate through itemized details
                for i in range(1, 4):  # Adjust the range based on the number of itemized sections you have
                    split = row[f'Split {i}']
                    item = row[f'Item {i}']
                    notes = row[f'Notes {i}']
                    
                    # Add each itemized detail to the email body
                    if pd.notnull(split) and pd.notnull(item) and pd.notnull(notes):
                        outlook_email.HTMLBody += f"- ${split} to {item} ({notes})<br>"
                    else:
                        break  # Exit loop if any of the fields is empty (assuming all are filled sequentially)

            else:
                outlook_email.HTMLBody += f"{date}, ${expense_amount}, Changed {expense_item_old} to {expense_correction} ({expense_notes})<br>"
        
        outlook_email.HTMLBody += f'<br><b>@{employee}, No action needed on your end. This is an email to inform you of the changes made.</b><br><br>'
        outlook_email.HTMLBody += f"<font color='red'><b><i>@{manager}, please check your Workday inbox for {expense_report}, review and approve. The Business Process requires a manager approval after changes are made. See changes made above.</i></b></font><br>"

        # Add your signature
        outlook_email.HTMLBody += signature.format(image_path='your_image_path_here')
        outlook_email.Display()
        break
        # outlook_email.Send()

def clean_and_filter(df):
    clean_filterd_df = df[df['Sent?'] == 'No']
    return clean_filterd_df 

outlook = initialize_outlook()
fixed_df = pd.read_excel(data_path, sheet_name='Fixed')
clean_filterd_df = clean_and_filter(fixed_df)
data = gather_corrections_data(clean_filterd_df)
email = create_email(outlook, data)

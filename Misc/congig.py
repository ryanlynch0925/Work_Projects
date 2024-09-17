import win32com.client as win32
import pandas as pd
from datetime import datetime
import traceback

try:
    outlook = win32.Dispatch('Outlook.Application')
except Exception as e:
    print(f"Error occurred while connecting to Outlook: {e}")
    traceback.print_exc()  # Print detailed traceback for debugging

excel_file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Master Expense Report V4.0.xlsx"
df = pd.read_excel(excel_file, sheet_name='Fixed')
filtered_df = df[df['Sent?'] == 'No']

# Create the Outlook application object
outlook = win32.Dispatch('Outlook.Application')

signature = '''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br>
            <img src="{image_path}" alt="Company Logo" style="width: 10px; height: auto;">
            '''

batch_emails = []
for index, row in filtered_df.iterrows():
    employee = row['Employee']
    expense_report = row['Expense Report']
    sent = row['Sent?']
    email = row['Email']
    manager = row['Manager']
    amount = row['Amount']
    date = row['Date']
    old = row['Old']
    new = row['New']
    notes = row['Notes']
    manager_email = row['Manager Email']

unique_employees = filtered_df.groupby(['Employee', 'Expense Report'])

for (employee, exp_report), group in unique_employees:
    email = group.iloc[0]['Email']
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
        outlook_email.CC = (f'Karla Kendrick; {manager_email}')
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
    
    # Send the email or store it in the batch_emails list for sending later
    # outlook_email.Send()  # Uncomment this line to send the email
    batch_emails.append(outlook_email)
    
# Send the batch emails
for email in batch_emails:
    email.Display()
    email.Send()
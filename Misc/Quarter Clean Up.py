import win32com.client as win32
import pandas as pd
import time

start_time = time.time()

#excel_file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\1.xlsx"
excel_file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Master Expense Report V4.0.xlsx"
due_date = '12/6/2023'
due_time = '5:00 PM'
quarter = 'Q3'

# Selecting specific columns
selected_columns = ['Employee', 'Credit Card Transaction', 'Amount', 'Load Date', 'Email', 'Sequence', 'Manager']
df = pd.read_excel(excel_file, sheet_name='Q3', usecols=selected_columns)

df['Load Date'] = df['Load Date'].dt.strftime('%m/%d/%Y')

# Convert 'Amount' column to numeric (remove '$' and commas)
df['Amount'] = df['Amount'].replace({'\$': '', ',': ''}, regex=True).astype(float)

# Perform aggregation
grouped_data = df.groupby('Employee').agg({
    'Credit Card Transaction': 'count',
    'Amount': 'sum',
    'Sequence': 'max',  # Take the max sequence for each employee
    'Email': 'first',   # Consider taking the first email of an employee
    'Manager': 'first'  # Consider taking the first manager of an employee
}).reset_index()

# Create the Outlook application object
outlook = win32.Dispatch('Outlook.Application')

# Function to append appropriate suffix to sequence numbers
def add_suffix(seq):
    if seq == 1:
        return '2nd'
    elif seq == 2:
        return '3rd'
    else:
        return f"{seq}th"

for index, row in grouped_data.iterrows():
    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.To = row['Email']
    mail.CC = 'Karla Kendrick' if row['Sequence'] == 0 else f"Karla Kendrick; {row['Manager']}"

    if row['Sequence'] > 0:
        mail.Subject = f"{quarter} Expenses Summary ({add_suffix(row['Sequence'])} Email)"
    else:
        mail.Subject = "{quarter} Expenses Summary"

    email_body = f"Dear {row['Employee']},<br><br>"
    email_body += f"You have {row['Credit Card Transaction']} {quarter} expense(s) totaling ${row['Amount']:,.2f}.<br><br>"
    email_body += f"<b><font color='red'>Kindly ensure that these expenses are submitted by {due_time} on {due_date}.</font></b><br>"

    employee_expenses = df[df['Employee'] == row['Employee']]
    email_body += "<br>Here is the list of your outstanding expenses:<br><br>"
    for _, expense in employee_expenses.iterrows():
        email_body += f"&emsp;<b>Transaction:</b> {expense['Credit Card Transaction']}<br>"
        email_body += f"&emsp;<b>Amount:</b> {expense['Amount']}<br>"
        email_body += f"&emsp;<b>Load Date:</b> {expense['Load Date']}<br>"

    signature = '''
    <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
    PH: +1 706-481-2635<br>
    T & E Specialist<br>
    Home Office<br>
    '''
    email_body += signature
    mail.HTMLBody = f"<html><body>{email_body}</body></html>"
    mail.Display()
    #break  # To send only the first email, remove this line to send all emails

# Record the end time
end_time = time.time()

# Calculate the elapsed time
elapsed_time = end_time - start_time

print(f"Total execution time: {elapsed_time:.2f} seconds")

import openpyxl
from openpyxl import load_workbook
import win32com.client as win32
import pandas as pd
import traceback
from datetime import datetime

try:
    outlook = win32.Dispatch('Outlook.Application')
except Exception as e:
    print(f"Error occurred while connecting to Outlook: {e}")
    traceback.print_exc()  # Print detailed traceback for debugging

file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\30-Day Report\Test.xlsx"
date = '01-15-2024'
df = pd.read_excel(file, sheet_name=date)
columns_to_remove = ['Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7']
df = df.drop(columns=columns_to_remove)

filtered_df = df[df['Rank'] <= 10]

signature = '''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br>
            <img src="{image_path}" alt="Company Logo" style="width: 10px; height: auto;">
            '''


# Create email body
email_body = '''I would like to draw your attention to an important matter pertaining to overdue expenses within our organization. Enclosed with this email is a PDF document detailing the list of Credit Card holders with expenses that are 30 days or older and have not been submitted.
'''
email_body += '''
Additionally, below, I have listed the top 10 employees along with the total amount of all outstanding expenses and the number of transactions for your reference:
'''
for index, row in filtered_df.iterrows():
    employee = row['Employee']
    total = row['Total']
    number_of_transactions = row['Number of Transactions']
    
    email_body += f"--{employee} with a total of ${total:,.2f} and {number_of_transactions} transactions.\n"

email_body += '''
Prompt attention to these overdue expenses is essential for accurate financial reporting and compliance purposes. 
I kindly request your immediate review of the attached document. Should you have any inquiries or require further clarification, please do not hesitate to contact me.
                '''

pdf_file_path = rf"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\30-Day Report\{date}.pdf"

# Create email
mail = outlook.CreateItem(0)
mail.Subject = "30-Days or older Report for Expenses"
mail.Body = email_body
mail.HTMLBody += signature

# Add recipients (modify the email addresses accordingly)
mail.CC = "tim@twavelead.com; bruce.maxwell@tidalwaveautospa.com"
mail.To = 'Leigh Stalling; Kevin McGonigle; Travis Powell; Karla Kendrick'

attachment = mail.Attachments.Add(pdf_file_path)

# Display the email (useful for testing)
mail.Display()

# Send the email (uncomment the line below when you are ready to send)
# mail.Send()
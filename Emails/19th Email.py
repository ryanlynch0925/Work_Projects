import win32com.client as win32
import pandas as pd
import time
from datetime import datetime

start_time = time.time()

excel_file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Master Expense Report V4.0.xlsx"
df = pd.read_excel(excel_file, sheet_name='Summary', header=3)
condensed_df = df[['Employee Name', '19th', 'Email', 'Manager']]
filtered_df = condensed_df[condensed_df['19th'] > 0]
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
    employee = row['Employee Name']
    email = row['Email']
    total = row['19th']
    manager = row['Manager']
    if total > 0:
            # Create the email
            mail = outlook.CreateItem(0)  # 0 represents an email item
                # Set the email properties
            mail.To = email
            #mail.CC = f"{manager}"
            mail.Subject = "Catch the Happy Wave: Expense Report Reminder! 🌊"
            #mail.Subject = "Testing Automatic Oustanding Emails"
            emailBody = f"Dear {employee},<br><br>" + \
            f'''
            ***If you have already recevied a reminder email, this is the updated expense total and should match your workday total.***<br><br>

            I hope you're ready to catch the <b style= 'color: #5BC2E7; font-size: 14pt'>Happy Wave</b> today! 🌞<br><br>

            This is a friendly reminder that the expense report is due <b style= 'color: #E76F44; font-size: 14pt'>20th of this month</b>
            Your expenses Total: <b style= 'color: #E76F44; font-size: 14pt'>${total:,.2f}</b>.<br><br>

            Got any questions or need some guidance to catch the <b style= 'color: #ED2891; font-size: 14pt'>Happy Expense Wave?</b> Just drop me a line! 🏄‍♂️<br><br>

            Thanks for being part of our <b style= 'color: #684199; font-size: 14pt'>Happy Wave!</b><br><br>
            '''
            emailBody += signature
            mail.HTMLBody = f"<html><body>{emailBody}</body></html>"
            batch_emails.append(mail)
    else:
          pass
for email in batch_emails:
    email.Display()
    #break
    email.Send()
# Record the end time
end_time = time.time()

    # Calculate the elapsed time
elapsed_time = end_time - start_time

print(f"Total execution time: {elapsed_time:.2f} seconds")
import pandas as pd
from Emails.config import initialize_outlook, signature
from datetime import datetime
import win32com.client as win32

def clean_and_filter(df):
    filtered_df = df[df['Sent'] == 'No']
    return filtered_df

def gather_corrections_data(clean_filterd_df):
    unique_employees = clean_filterd_df.groupby(['Employee', 'Consultant'])
    return unique_employees

def create_email(outlook, unique_employees):
    
    for (employee, consultant), group in unique_employees:
        consultant = group.iloc[0]['Consultant']
        employee = group.iloc[0]['Employee']
        amount = group.iloc[0]['Amount']
        quantity = group.iloc[0]['# of GCs']

        gift_card_amounts = []
        for i in range(1,10):
            gift_card_amount = group.iloc[0][f'{i}']

            if gift_card_amount is not None:
                gift_card_amounts.append(gift_card_amount)
            else:
                break


        email = outlook.CreateItem(0)  # Create an Outlook email object
        
        # Set up the email content
        email.To = consultant
        email.CC = 'Keri Pack'
        email.Subject = "Gift Cards Purchased with Company Credit Card"
        email.HTMLBody += f"Dear {consultant},<br><br>"
        email.HTMLBody += f"{employee}'s company credit card was used to buy gift cards.<br><br>"
        email.HTMLBody += f"{employee} spent ${amount:,.2f} on {quantity} gift cards for the following amounts:<br>"
        for gift_card_amount in gift_card_amounts:
            if pd.notna(gift_card_amount):
                email.HTMLBody += f"--${gift_card_amount}<br>"
        email.HTMLBody += f'''<br>
        The company credit card should not be used to purchase gift cards for team members. All monetary rewards will go through payroll, please see attached documentation for further instructions.<br>
        '''
        email.HTMLBody += signature
        pdf_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\.000001 TW Expenses\Policy\Monetary Award Policy.pdf"
        email.Attachments.Add(pdf_path)
        email.Display()
        # break
        email.Send()


data_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Corrections Template.xlsx"
outlook = initialize_outlook()
df = pd.read_excel(data_path, sheet_name = "Gift Cards")
clean_filterd_df = clean_and_filter(df)
data = gather_corrections_data(clean_filterd_df)
email = create_email(outlook, data)
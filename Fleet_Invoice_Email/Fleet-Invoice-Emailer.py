import os
import re
import PyPDF2
import win32com.client as win32
from datetime import datetime
import time

start_time = time.time()
image_path = r"C:\Users\MistyDouglas\OneDrive - Tidal Wave Autospa\Desktop\Company Logo.png"
####################### Only Change this line ##########################
month = 'December'
month_folder = 'DEC 2023 SALES AUTO'
##################### Only if needed to be changed #####################
signature = f'''
    <br><span style="font-family:'Bradley Hand ITC', cursive, sans-serif; color: #0C1731; font-size: 16pt;">Misty Douglas<br></span>
    <i>Accounts Receivable (Fleet)</i><br><br>

    PO Box 311<br>
    Thomaston, GA 30286<br>
    O: 706-647-0414 x146<br>
    A: 706-535-2911<br>
    <img src="{image_path}" alt="Company Logo" style="width: 10px; height: auto;">
    '''
folder_path = rf"C:\Users\MistyDouglas\OneDrive - Tidal Wave Autospa\Documents\ACCOUNTS RECEIVABLE\FLEET\HOME OFFICE\PDFs to EMAIL\{month_folder}"
exclude_domain = "fleetbilling@tidalwaveautospa.com"
email_pattern = r'\b[A-Za-z0-9._%+-]+[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
##########################################################################

def extract_customer_info(pdf_text):
    pass

def extract_emails(pdf_path):
    pdf_text = ""
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            pdf_text += page.extract_text()
    
    emails = re.findall(email_pattern, pdf_text)
    filtered_emails = [email for email in emails if exclude_domain not in email]
    return filtered_emails

def extract_past_due_info(pdf_text):
    pattern = r'\$([\d,.]+)\s+\$([\d,.]+)\s+\$([\d,.]+)\s+\$([\d,.]+)\s+\$([\d,.]+)\s+\$([\d,.]+)'
    matches = re.search(pattern, pdf_text)
    if matches:
        current_due = matches.group(1)
        past_due_1_30 = matches.group(2)
        past_due_31_60 = matches.group(3)
        past_due_61_90 = matches.group(4)
        past_due_90_plus = matches.group(5)
        balance = matches.group(6)
        
        return {
            "Current Due": current_due,
            "Past Due 1-30 Days": past_due_1_30, 
            "Past Due 31-60 Days": past_due_31_60,
            "Past Due 61-90 Days": past_due_61_90,
            "Past Due 90+ Days": past_due_90_plus,
            "Balance": balance
        }
    else:
        return None
    
for root, dirs, files in os.walk(folder_path):
    for filename in files:
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(root, filename)
            #print(pdf_path)
            extracted_emails = extract_emails(pdf_path)
            
            with open(pdf_path, "rb") as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                pdf_text = ""
                for page in pdf_reader.pages:
                    pdf_text += page.extract_text()
                    #print(pdf_text)
                    account_info = extract_past_due_info(pdf_text)
                    #print(account_info)
                    email_body = ''
                    # No Past Dues
                    if (
                        float(account_info["Past Due 1-30 Days"].replace(",","")) == 0 and 
                        float(account_info["Past Due 31-60 Days"].replace(",","")) == 0 and 
                        float(account_info["Past Due 61-90 Days"].replace(",","")) == 0 and 
                        float(account_info["Past Due 90+ Days"].replace(",","")) == 0 and 
                        float(account_info["Balance"].replace(",","")) != 0
                    ):
                        email_body = '''

                        Thank you for choosing Tidal Wave Fleet! I am attaching your Tidal Wave Auto Spa fleet invoice above. 
                        To ensure timely processing of your payment, please submit all checks to the home office address noted on the invoice.<br>
                        If you wish to pay by credit card, please click here: 
                        <a href="https://tidalwavefleet.securepayments.cardpointe.com/pay" style="background-color: #003896; color: #FFFFFF; text-decoration: none; padding: 10px 20px; border-radius: 20px; display: inline-block">Pay by Credit Card</a></p>
                        <br>
                        <strong>Please note:</strong><br>
                        <ul>
                        <li>Due to the centralization of our fleet billing system, we kindly request that payments <b>(both check and credit card)</b> no longer be made at the site locations.<br></li>
                        <li>Aging detail is now being printed on the invoice but may not reflect a recent payment.<br></li>
                        </ul>

                        Please feel free to respond to this email with any changes/corrections to your information. For instance, if you would like this invoice emailed to a different email address or wish to correct a phone number. Also, please don’t hesitate to reach out if you have any questions.</li><br><br>
                        
                        Keep your <b style= "color: #0085CA">Fleet</b> Neat!<br>
                        '''
                    elif (
                        float(account_info["Past Due 1-30 Days"].replace(",","")) > 0 or
                        float(account_info["Past Due 31-60 Days"].replace(",","")) > 0 or 
                        float(account_info["Past Due 61-90 Days"].replace(",","")) > 0 or
                        float(account_info["Past Due 90+ Days"].replace(",","")) > 0 and 
                        float(account_info["Balance"].replace(",","")) != 0
                    ):
                        email_body = f'''

                        Thank you for choosing Tidal Wave Fleet! I am attaching your Tidal Wave Auto Spa fleet invoice above. 
                        To ensure timely processing of your payment, please submit all checks to the home office address noted on the invoice.<br>
                        If you wish to pay by credit card, please click here: 
                        <a href="https://tidalwavefleet.securepayments.cardpointe.com/pay" style="background-color: #003896; color: #FFFFFF; text-decoration: none; padding: 10px 20px; border-radius: 20px; display: inline-block">Pay by Credit Card</a></p>
                        <br>
                        <span style="color: red;">Reminder – We are still awaiting payment(s) for past due invoice(s). If not already sent, please feel free to make one payment of <b>${account_info['Balance']}</b> for the total account balance.<br><br></span>
                        <strong>Please note:</strong><br>
                        <ul>
                        <li>Due to the centralization of our fleet billing system, we kindly request that payments <b>(both check and credit card)</b> no longer be made at the site locations.<br></li>
                        <li>Aging detail is now being printed on the invoice but may not reflect a recent payment.<br></li>
                        </ul>

                        Please feel free to respond to this email with any changes/corrections to your information. For instance, if you would like this invoice emailed to a different email address or wish to correct a phone number. Also, please don’t hesitate to reach out if you have any questions.</li><br><br>
                        
                        Keep your <b style= "color: #0085CA">Fleet</b> Neat!<br>
                        '''
                email_body += signature
                # Create the Outlook application object
                outlook = win32.Dispatch('Outlook.Application')

                mail = outlook.CreateItem(0)
                mail.SentOnBehalfOfName = exclude_domain
                mail.To = '; '.join(extracted_emails)  # Join the emails with a comma and space
                mail.CC = ''
                mail.Subject = f'Tidal Wave Auto Spa Fleet Invoice - {month}'
                mail.HTMLBODY = f"<html><body>{email_body}</body></html>"

                mail.Attachments.Add(pdf_path)
                mail.Display()
                break
                #mail.Send()

end_time = time.time()

elapsed_time = end_time - start_time

print(f"Total execution time: {elapsed_time:.2f} seconds")


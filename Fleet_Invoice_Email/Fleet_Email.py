import subprocess
import sys
import os
import string

### Required Packages ###
def clean_string(s):
    # Create a whitelist of printable ASCII characters
    printable = set(string.printable)
    
    # Remove non-printable characters
    cleaned = ''.join(filter(lambda x: x in printable, s)).strip()
    
    # Remove any leading or trailing spaces
    cleaned = cleaned.strip()
    
    # Return the cleaned string
    return cleaned

def install_required_packages(requirement_file):
    with open(requirement_file, 'r') as file:
        required_packages = [clean_string(line.split('==')[0]) for line in file.readlines() if line.strip()]

    for package in required_packages:
        if package:  # Ensure the package name is not an empty string
            print(f"Checking package: '{package}'")
            if not is_package_installed(package):
                print(f"Installing {package}...")
                try:
                    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                    print(f"{package} installed successfully.")
                except Exception as e:
                    print(f"Error installing package {package}: {e}")
                    raise

def is_package_installed(required_package):
    try:
        subprocess.check_output([sys.executable, '-m', 'pip','show', required_package])
        return True
    except subprocess.CalledProcessError:
        return False
requirements_file = os.path.join(os.path.dirname(__file__),'requirements.txt')
install_required_packages(requirements_file)

### Main Code ###
import re
import PyPDF2
import win32com.client as win32
from datetime import datetime
import timeit
import logging
from fleet_constants import email_pattern, exclude_domain, signature, month, folder_path

logging.basicConfig(filename='pdf_processing_errors.log', level=logging.ERROR, format='%(asctime)s - %(message)s')

def extract_emails(pdf_text):
    try:
        emails = re.findall(email_pattern, pdf_text)
        filtered_emails = [email for email in emails if exclude_domain not in email]
        return filtered_emails
    except Exception as e:
        error_message = f"Error extracting emails: {e}"
        logging.error(error_message)
        return []

def extract_past_due_info(pdf_text):
    try:
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
            return {}
    except Exception as e:
        error_message = f"Error extracting past due info: {e}"
        logging.error(error_message)
        return {} 

def read_text(pdf_path):
    try:
        with open(pdf_path, "rb") as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            pdf_text = ""
            for page_number in range(2):  # Loop through the first two pages
                pdf_text += pdf_reader.pages[page_number].extract_text().strip()
        return pdf_text
    except PyPDF2.utils.PdfReadError as e:
        error_message = f"Error reading PDF {pdf_path}: {e}"
        logging.error(error_message)
        return ''
    except Exception as e:
        error_message = f"Unexpected error reading PDF {pdf_path}: {e}"
        logging.error(error_message)
        return ''
def create_email(account_info, extracted_emails, pdf_path):
    try:
        body = ''

        # No Past Dues
        if all(float(account_info[key].replace(",", "").strip()) == 0 for key in account_info.keys() if key != "Balance" and key != "Current Due"):
            body = '''
            Thank you for choosing Tidal Wave Fleet! I am attaching your Tidal Wave Auto Spa fleet invoice above. 
            To ensure timely processing of your payment, please submit all checks to the home office address noted on the invoice.<br>
            If you wish to pay by credit card, please click here: 
            <a href="https://tidalwavefleet.securepayments.cardpointe.com/pay" style="background-color: #003896 ; color: #FFFFFF; text-decoration: none; padding: 10px 20px; border-radius: 20px; display: inline-block">Pay by Credit Card</a></p>
            <br>
            <strong>Please note:</strong><br>
            <ul>
            <li>Due to the centralization of our fleet billing system, we kindly request that payments <b>(both check and credit card)</b> no longer be made at the site locations.<br></li>
            <li>Aging detail is now being printed on the invoice but may not reflect a recent payment.<br></li>
            </ul>

            <b span style="text-decoration: underline">Fleet House Accounts using Fleet Cards</b><br><br>
            <ul>
            <li>Effective immediately, House Accounts using Fleet Cards will not select a wash type when they approach the XPT. The driver will simply scan their Fleet Card and the wash that the account is set up to redeem will load. If the driver selects a wash that is not the correct wash for their account, the transaction will cancel out, and they will need to scan their Fleet Card first to complete a transaction with the appropriate wash.<br>
            
            <li>If the customer is an “Any Wash” House Account customer using a Fleet Card, they will select their wash and then scan their card when the wash is ready to tender.</li>
            </ul>

            Please feel free to respond to this email with any changes/corrections to your information. For instance, if you would like this invoice emailed to a different email address or wish to correct a phone number. Also, please don’t hesitate to reach out if you have any questions.</li><br><br>
            
            Keep your <b style= "color: #0085CA">Fleet</b> Neat!<br>
            '''
        ### Past Dues ###
        else:
            body = f'''
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

            <b span style="text-decoration: underline">Fleet House Accounts using Fleet Cards</b><br><br>
            <ul>
            <li>Effective immediately, House Accounts using Fleet Cards will not select a wash type when they approach the XPT. The driver will simply scan their Fleet Card and the wash that the account is set up to redeem will load. If the driver selects a wash that is not the correct wash for their account, the transaction will cancel out, and they will need to scan their Fleet Card first to complete a transaction with the appropriate wash.<br>
            
            <li>If the customer is an “Any Wash” House Account customer using a Fleet Card, they will select their wash and then scan their card when the wash is ready to tender.</li>
            </ul>

            Please feel free to respond to this email with any changes/corrections to your information. For instance, if you would like this invoice emailed to a different email address or wish to correct a phone number. Also, please don’t hesitate to reach out if you have any questions.</li><br><br>
            
            Keep your <b style= "color: #0085CA">Fleet</b> Neat!<br>
            '''

        body += signature
        # Create the Outlook application object
        outlook = win32.Dispatch('Outlook.Application')
        outlook_email = outlook.CreateItem(0)
        outlook_email.SentOnBehalfOfName = exclude_domain
        outlook_email.To = '; '.join(extracted_emails)  # Join the emails with a comma and space
        outlook_email.CC = ''
        outlook_email.Subject = f'Tidal Wave Auto Spa Fleet Invoice - {month}'
        outlook_email.HTMLBODY = f"<html><body>{body}</body></html>"
        outlook_email.Attachments.Add(pdf_path)
        return  outlook_email
    except Exception as e:
        error_message = f"Error creating email: {e}"
        logging.error(error_message)
        return None
    
def send_emails(emails_to_send):
    if emails_to_send:
        for email in emails_to_send:
            email.Display()
            # email.Send()
    else:
        print("No emails to send!")

def main():
    start_time = timeit.default_timer()
    emails_to_send = []
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(root, filename)
                pdf_text = read_text(pdf_path)
                extracted_emails = extract_emails(pdf_text)
                account_info = extract_past_due_info(pdf_text)
                print(account_info.values())
                if account_info is not None:
                    email = create_email(account_info, extracted_emails, pdf_path)
                    if email is not None:
                        emails_to_send.append(email)
                else:
                    print(f"Skipping email creation for {pdf_path} due to error in extracting past due info.")

    send_emails(emails_to_send)
    elapsed_time = timeit.default_timer() - start_time
    print(f"Elapsed time: {round(elapsed_time, 3) } seconds")

if __name__ == "__main__":
    main()
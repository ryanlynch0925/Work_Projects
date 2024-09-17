import traceback
import win32com.client as win32
import pandas as pd

data_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Corrections Template.xlsx"
image_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Scripts\Company Logo.png"
signature = f'''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br><br>
            <img src="{image_path}" alt="Company Logo">
            '''

### Top 40 Email ###
top_40_CC = "Leigh Stalling; Kevin McGonigle; Travis Powell; Karla Kendrick;"
top_40_sheet_name = 'Summary'
top_40_header = 7

### Reports Fixed Email ###
fixed_CC = 'Karla Kendrick'
fixed_sheet_name = 'Fixed'
cheatsheet_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\.000001 TW Expenses\Expense Item Cheat Sheat.docx"
itemization_instructions_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\.1 Training\Itemizing Receipts.docx"
amazon_invoice_error = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\PDF Instructions\Amazon_Order_Instructions.pdf"
walmart_dot_com_invoice_error = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\PDF Instructions\Walmart_Order_Instructions.pdf"
globalIndustrial = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\.1 Training\GlobalIndustrial.pdf"

### Sent Back or Removed Email ###
sent_back_CC = 'Karla Kendrick'
sent_back_sheet_name = 'Corrections'
sent_back_subject = "Expense Report Sent Back- Errors"
removed_subject = "An Expense Transaction has been Removed from your Expense Report"
cancelled_subject = "An Expense Report has been Cancelled"
correction_notes = {
    "Incomplete Invoice": "Please attach the complete invoice.",
    "Reimbursement": "Please verify if this is a reimbursement or a company expense.",
    "Report Canceled": "The report has been canceled because the receipts are not linked to a credit card transaction.",
    "Incorrect Format": "Please attach invoices in PDF, JPEG, or PNG format.",
    "Non-Matching Amount": "Please attach an invoice that matches the charge amount.",
    "Poor Quality": "The invoice is unreadable; please retake the image and re-upload it.",
    "Lost Receipt Form": "Please fill out and attach Lost Receipt Form.",
    "Expense Item Miscoding": "Please update the expense item to accurately reflect the purchased items.",
    "Location Miscoding": "Please update the location to accurately reflect where the charge was incurred.",
    "Contact Merchant": "Please contact the Merchant to obtain the invoice. Expense Amount is too large for Lost Receipt Form.",
    "Wrong Receipt": "Please attach an invoice that matches the total amount, date, and Merchant.",
    "Memo Error": "Please fill out the memo with details of what was purchased.",
    "Attendees Missing": "Please list all attendees who were present.",
    "Non-SL Form Required": "Please complete the Non-SL form and attach it to the transaction. Ensure it is approved by Tim, Bruce, or the Consultant.",
    "Personal Expense?": "Is this a personal expense? If not, please uncheck the personal expense box.",
    "Receipt Link Error": "Receipt is not linked to a credit card transaction.",
    "Mileage Reimbursement": "Please send mileage reimbursements to payroll.",
    "Attachment Load Error": "Please send the invoice to expenses@twavelead.com due to an error loading the invoice in Workday.",
    "Separate Receipts": "Please separate each charge into its own line item.",
    "Amazon Invoice Error": "Please refer to the attached Amazon Instructions to resolve this correction.",
    "Walmart.com Invoice Error": "Please refer to the attached Walmart.com Instructions to resolve this correction.",
    "Itemization Needed": "Please refer to the attached cheat sheet to itemize the expense line correctly.",
    "Missing Receipt": "Please ensure the 'Missing Receipt?' checkbox is checked for any missing receipts.",
    "Recheck Itemization": "Please review the itemization; one or more lines are incorrect and need to be revised.",
    "Global Industrial": "Please refer to the attached Global Industrial Instructions for future purchases."
}

### Resuable Functions ###
def initialize_outlook():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        return outlook
    except Exception as e:
        print(f'Error occured while initializing Outlook: {e}')
        traceback.print_exc()

def gather_corrections_data(clean_filterd_df):
    unique_employees = clean_filterd_df.groupby(['Employee', 'Email', 'Expense Report'])
    return unique_employees
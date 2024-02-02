import traceback
import win32com.client as win32

data_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Master Expense Report V4.0.xlsx"
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

### Sent Back or Removed Email ###
sent_back_CC = 'Karla Kendrick'
sent_back_sheet_name = 'Corrections'
sent_back_subject = "An Expense Report has been Sent Back to you for Corrections"
removed_subject = "An Expense Transaction has been Removed from your Expense Report"

### Resuable Functions ###
def initialize_outlook():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        return outlook
    except Exception as e:
        print(f'Error occured while initializing Outlook: {e}')
        traceback.print_exc()

def gather_corrections_data(clean_filterd_df):
    unique_employees = clean_filterd_df.groupby(['Employee', 'Email'])
    return unique_employees
import pandas as pd
from datetime import datetime

# File and Sheet Variables
file_name = 'Personal Expenses.xlsx'
excel_file = rf"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Personal Reimbursements\{file_name}"
sheet_name = input("What is the sheet name?")
print(sheet_name)
image_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Scripts\Company Logo.png"

signature = f'''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br><br>
            <img src="{image_path}" alt="Company Logo">
            '''

df = pd.read_excel(excel_file, sheet_name=sheet_name)
df = df[['Employee', 'Total', 'Email']]
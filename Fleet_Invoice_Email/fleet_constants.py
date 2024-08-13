### Only Change this line ###
month = 'September'
month_folder = 'AUGUST 2023 SALES'
image_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Invoices\Misty\Company Logo.png"
signature = f'''
    <br><span style="font-family:'Bradley Hand ITC', cursive, sans-serif; color: #0C1731; font-size: 16pt;">Misty Douglas<br></span>
    <i>Accounts Receivable (Fleet)</i><br><br>

    PO Box 311<br>
    Thomaston, GA 30286<br>
    O: 706-647-0414 x146<br>
    A: 706-535-2911<br>
    <img src="{image_path}" alt="Company Logo" style="width: 10px; height: auto;">
    '''

# folder_path = rf"C:\Users\MistyDouglas\OneDrive - Tidal Wave Autospa\Documents\ACCOUNTS RECEIVABLE\FLEET\HOME OFFICE\PDFs to EMAIL\{month_folder}"
folder_path = rf'C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Scripts\PyTest'
exclude_domain = "fleetbilling@tidalwaveautospa.com"
email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
##################################
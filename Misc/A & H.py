import os
import re
import PyPDF2
import time
import shutil
import tempfile
import subprocess

folder_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Invoices\A & H Unproccesed"
amount_pattern = r'\$[0-9,]+\.\d{2}'  # Pattern to match $XXX.XX but not $0.00
invoice_pattern = r'Invoice#(\d+)'  # Pattern to match Invoice# followed by a number
invoice_date_pattern = r'\d{1,2}/\d{1,2}/\d{4}'


for root, dirs, files in os.walk(folder_path):
    for filename in files:
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(root, filename)
            
            with open(pdf_path, "rb") as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                pdf_text = ""
                for page in pdf_reader.pages:
                    pdf_text += page.extract_text()

                # Remove white spaces (spaces, tabs, and line breaks) from the text
                pdf_text_stripped = re.sub(r'\s', '', pdf_text)
                #print(pdf_text_stripped)

                # Use regular expressions to find the first occurrence of the amount pattern
                amount_match = re.search(amount_pattern, pdf_text_stripped)
                invoice_match = re.search(invoice_pattern, pdf_text_stripped)
                date_match = re.search(invoice_date_pattern, pdf_text_stripped)

                if amount_match and invoice_match and date_match:
                    amount = amount_match.group()  # Extract the amount
                    invoice_number = invoice_match.group(1)  # Extract the number after the #
                    date = date_match.group(0)  # Extract the date

                    formatted_date = date.replace('/', '_')

                    time.sleep(1)

                    # Create a new PDF file path with the desired name
                    new_filename = f"{formatted_date}, {amount}({invoice_number}).pdf"
                    new_pdf_path = os.path.join(root, new_filename)

                    # Copy the original PDF to the new location
                    shutil.copy(pdf_path, new_pdf_path)

                    # Delete the original PDF using a scheduled task
                    try:
                        temp_script = os.path.join(tempfile.gettempdir(), "delete_file.py")
                        with open(temp_script, "w") as script_file:
                            script_file.write(f'import os\nos.remove("{pdf_path}")')

                        # Schedule a task to delete the file on the next reboot
                        subprocess.run(['schtasks', '/create', '/tn', 'delete_pdf_task', '/sc', 'onstart', '/tr', f'pythonw "{temp_script}"'])
                    except Exception as e:
                        print(f"Error scheduling file deletion: {e}")

                    print(f"Renamed to: {new_pdf_path}")

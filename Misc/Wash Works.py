import re
import PyPDF2
import os

def extract_payment_amount(text):
    patterns = [
        r'deposits\s*to\s*de\s*te\s*of([\d.,\s]+)',
        r'deposits\s*to\s*da\s*te\s*of\s*([\d.\s]+)',
        r'deposits\s*to\s*<la\s*te\s*of\s*([\d.\s]+)',
        r'deposits\s*to\s*dc:te\s*of\s*([\d.\s]+)',
        r'deposits\s*to\s*dcte\s*of\s*([\d.\s]+)',
        r'deposits\s*to\s*de:\s*te\s*of\s*([\d.\s]+)'
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
                amount_str = match.group(1).replace(' ','').strip()
                amount_str = amount_str.rstrip('.')
                amount = float(amount_str)
                return amount
    return None

def extract_invoice_number(text):
    patterns = [
        r'INVOICE\s+NO:\s*(\d+)',
        r'INVOICE\s+NO·\s*(\d+)',
        r'INVOICE\s+NO[-·:\s]*\.?(\d+)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
        return None  

def main():
    input_pdf_path = r'C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Invoices\Wash Works\all.pdf'

    # Get the directory path of the input PDF file
    input_dir = os.path.dirname(input_pdf_path)

    input_pdf = open(input_pdf_path, 'rb')
    pdf_reader = PyPDF2.PdfReader(input_pdf)

    pages_with_no_payment = []

    for page_num, page in enumerate(pdf_reader.pages, start=1):
        page_text = page.extract_text()
        payment_amount = extract_payment_amount(page_text)
        invoice_number = extract_invoice_number(page_text)
        if payment_amount and invoice_number:
            #print(f'Page {page_num}, Invoice Number: ({invoice_number}) Payment Amount: ${payment_amount:.2f}')
            # Generate the new file name using the extracted amount
            new_pdf_name = f'{payment_amount} ({invoice_number}).pdf'
            
            new_pdf_path = os.path.join(input_dir, new_pdf_name)

            # Create a PDF writer
            output_pdf = PyPDF2.PdfWriter()
            output_pdf.add_page(page)
            
            # Save the extracted page as a new PDF with the new name
            with open(new_pdf_path, 'wb') as output_file:
                output_pdf.write(output_file)
            
            print(f'Page {page_num} saved as: {new_pdf_name}')
        elif invoice_number or payment_amount:
            if payment_amount:
                print(f'Page {page_num}, Payment Amount: ${payment_amount:.2f}')
                #print(page_text)
                pages_with_no_payment.append(page_num)
                
            if invoice_number:
                print(f'Page {page_num}, Invoice Number: {invoice_number}')
                #print(page_text)
                pages_with_no_payment.append(page_num)
            
        else:
            pages_with_no_payment.append(page_num)
            #print(page_text)
    input_pdf.close()

    if pages_with_no_payment:
        print("\nSummary: Pages with no payment amount found:")
        for page_num in pages_with_no_payment:
            print(f'Page {page_num} has no payment amount or invoice number.')
    else:
        print("\nAll pages have a payment amount or invoice number.")


if __name__ == '__main__':
    main()


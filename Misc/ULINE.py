import re
import PyPDF2
import os

def extract_amount(text):
    text_without_spaces = text.replace(" ", "")
    #print(text_without_spaces)
    patterns = [
        r'CREDITCARD([\d,.]+)',
        r'CREDITCAl~D([\d,.]+)',
        r'CREDITCARD([\d,.]+)',
        r'CHARGEDTOVISAENDINGIN(\d+)\$(\d+\.\d{2})',
        r'CHARGEDTOVISAENDINGIN\n(\d+)\$(\d+\.\d{2})',
        r'CHARGEDTOVISAENDINGIN(\d+)(?:\n|\s)+\$([\d,.]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text_without_spaces, re.IGNORECASE)
        if match:
            # Return the matched amount or group based on the pattern
            if len(match.groups()) > 1:
                return match.group(2)
            else:
                return match.group(1)
    return None

def main():
    Zeroto199_99 = '0 to 199.99'
    TwoHundredto799_99 = '200 to 799.99'
    EightHundredto2499_99 = '800 to 2499.99'
    TwoThousandFiveHundredandUp = '2500 and Up'

    #input_pdf_path = rf"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Invoices\ULINE\{Zeroto199_99}\{Zeroto199_99}.pdf"  # Replace with your PDF file path
    #input_pdf_path = rf"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Invoices\ULINE\{TwoHundredto799_99}\{TwoHundredto799_99}.pdf"
    # input_pdf_path = rf"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Invoices\ULINE\{EightHundredto2499_99}\{EightHundredto2499_99}.pdf"
    input_pdf_path = rf"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Invoices\ULINE\Unprocessed.pdf"

    # Get the directory path of the input PDF file
    input_dir = os.path.dirname(input_pdf_path)
    
    # Open the PDF file in read-binary mode
    input_pdf = open(input_pdf_path, 'rb')
    pdf_reader = PyPDF2.PdfReader(input_pdf)

    for page_num, page in enumerate(pdf_reader.pages, start=1):
        # Extract text from the current page
        page_text = page.extract_text()
        # print(page_text)

        #print(f'Page {page_num} Text: {page_text}')
        # Extract the amount using the defined function
        amount = extract_amount(page_text)

        if amount:
            # Generate the new file name using the extracted amount
            new_pdf_name = f'{amount.replace("$", "").replace(",", "")}.pdf'
            
            new_pdf_path = os.path.join(input_dir, new_pdf_name)

            # Create a PDF writer
            output_pdf = PyPDF2.PdfWriter()
            output_pdf.add_page(page)
            
            # Save the extracted page as a new PDF with the new name
            with open(new_pdf_path, 'wb') as output_file:
                output_pdf.write(output_file)
            
            # print(f'Page {page_num} saved as: {new_pdf_name}')
        else:
            print(f'Amount not found on Page {page_num}')
    # Close the input PDF file
    input_pdf.close()

if __name__ == '__main__':
    main()

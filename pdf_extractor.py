import PyPDF2
#import pathlib
import sys
import os
import re
import openpyxl

def extract_tax_codes_from_pdf(pdf_path, start_page = 1):
    tax_codes = []
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        # Get the number of pages in the PDF
        num_pages = len(pdf_reader.pages)

        # Ensure the start_page is within the valid range
        start_page = max(min(start_page, num_pages), 1)

        # Iterate through each page
        for page_number in range(start_page, num_pages):
            # Get the page
            page = pdf_reader.pages[page_number]

            # Extract text from the page
            text = page.extract_text()

            matches = re.findall(r'(?<![a-zA-Z0-9])\d{10}(?:-\d{3})?(?![a-zA-Z0-9])', text)

            tax_codes.extend([match for match in matches if (match != '0300942001-022')])
            print(f"{tax_codes}")

            # Print the extracted text for the current page
            #print(f"Page {page_number + 1} text:\n{text}\n")
        pdf_file.close()
    return tax_codes
            

def get_pdf_files(directory):
    pdf_files = []

    # Iterate through all files in the directory
    for filename in os.listdir(directory):
        # Check if the file has a ".pdf" extension
        if filename.endswith(".pdf"):
            # Build the full path to the PDF file
            pdf_path = os.path.join(directory, filename)
            
            # Add the PDF file path to the list
            pdf_files.append(pdf_path)

    return pdf_files

def export_to_excel(tax_codes_matrix):
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Select the active sheet (default is the first sheet)
    sheet = workbook.active

    # Starting row to write the data
    start_row = 1

    # Loop through the list of lists and write it to consecutive rows and columns
    for row_index, tax_codes in enumerate(tax_codes_matrix):
        for col_index, tax_code in enumerate(tax_codes):
            sheet.cell(row=start_row + row_index, column=col_index + 1, value=tax_code)

    # Save the workbook to a file
    workbook.save('mst.xlsx')

    print('Data written to the Excel file within a loop successfully.')


# Path to the current script
#current_path = pathlib.Path(__file__).parent.resolve()

# Get the path to the executable
exe_path = sys.argv[0]

# Get the directory containing the executable
exe_dir = os.path.dirname(exe_path)

# Open a file in write mode ('w')
with open('log.txt', 'w') as file:
    # Write content to the file
    file.write(exe_dir)

# File is automatically closed when you exit the 'with' block

pdf_files = get_pdf_files(exe_dir)
tax_codes_matrix = []

# Extracting PDF files
for pdf_file in pdf_files:
    print(f"Extracting tax codes in: {pdf_file}")
    tax_codes_matrix.append(extract_tax_codes_from_pdf(pdf_file))

export_to_excel(tax_codes_matrix)



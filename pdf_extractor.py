import PyPDF2
import pathlib
import os
import re

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        # Get the number of pages in the PDF
        num_pages = len(pdf_reader.pages)

        # Iterate through each page
        for page_number in range(num_pages):
            # Get the page
            page = pdf_reader.pages[page_number]

            # Extract text from the page
            text = page.extract_text()

            taxCodes = re.findall(r'(?<![a-zA-Z0-9])\d{10}(?:-\d{3})?(?![a-zA-Z0-9])', text)
            print(f"{taxCodes}")

            # Print the extracted text for the current page
            #print(f"Page {page_number + 1} text:\n{text}\n")
            

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

# Path to the current script
current_path = pathlib.Path(__file__).parent.resolve()
pdf_files = get_pdf_files(current_path)


# Extracting PDF files
for pdf_file in pdf_files:
    print(f"Extracting: {pdf_file}")
    extract_text_from_pdf(pdf_file)

import pandas as pd
from docx import Document

def read_replacements_from_excel(excel_path):
    """Read replacements from Excel and return a dictionary."""
    df = pd.read_excel(excel_path)  # Read the Excel file

    # for index, row in df.iterrows():

    return dict(zip(df['Placeholder'], df['Value']))  # Convert to dictionary

def replace_placeholders(doc_path, output_path, replacements):
    """Replace placeholders in a Word document."""
    doc = Document(doc_path)
    
    # Replace placeholders in paragraphs
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    
    # Replace placeholders in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

    doc.save(output_path)
    print(f"File saved as {output_path}")

# File paths
excel_file = "data.xlsx"
word_template = "template.docx"
output_word = "filled_document.docx"

# Process the files
replacements = read_replacements_from_excel(excel_file)
replace_placeholders(word_template, output_word, replacements)

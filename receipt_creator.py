import pandas as pd
from docx import Document
# import shutil as sh

def generate_receipts_from_template_with_filled_data():
    data_file = "data.xlsx"
    template_file = "template.docx"

    """Read replacements from Excel and return a dictionary."""
    df = pd.read_excel(data_file)  # Read the Excel file

    # Define starting index
    start_index = 2

    for index, row in df.iloc[start_index:].iterrows():
        row_dict = {i: value for i, value in enumerate(df.iloc[index])}
        replace_placeholders(template_file, f"filled_document_{index}.docx", row_dict)

    # return dict(zip(df['Placeholder'], df['Value']))  # Convert to dictionary

def replace_placeholders(doc_file, output_file, replacements):
    # sh.copy(doc_path, output_path)

    """Replace placeholders in a Word document."""
    doc = Document(doc_file)
    
    # Replace placeholders in paragraphs
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            placeholder = f"{{{key}}}"
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, str(value))
    
    # Replace placeholders in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    placeholder = f"{{{key}}}"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, str(value))

    doc.save(output_file)
    print(f"File saved as {output_file}")

# File paths
# excel_file = "data.xlsx"
# word_template = "template.docx"
# output_word = "filled_document.docx"

# Process the files
# replacements = read_replacements_from_excel(excel_file)
# replace_placeholders(word_template, output_word, replacements)

generate_receipts_from_template_with_filled_data()

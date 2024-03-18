from pdf2image import convert_from_path
import pytesseract
from PIL import Image

def extract_text_from_scanned_pdf(pdf_path):
    # Convert PDF to images
    images = convert_from_path(pdf_path, 500)  # Adjust the DPI as needed

    # Extract text from each image using Tesseract OCR
    extracted_text = ""
    for i, image in enumerate(images):
        image_path = f"temp_image_{i}.png"
        image.save(image_path, 'PNG')
        text = pytesseract.image_to_string(Image.open(image_path))
        extracted_text += f"Page {i + 1} text:\n{text}\n\n"

    return extracted_text

# Replace "path/to/your/scanned.pdf" with the actual path to your scanned PDF
pdf_path = "scanned_tax_codes.pdf"
text_result = extract_text_from_scanned_pdf(pdf_path)
print(text_result)

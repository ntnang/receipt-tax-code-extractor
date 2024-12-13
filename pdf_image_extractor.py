from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import cv2
import numpy as np

def extract_text_from_scanned_pdf(pdf_path):
    # Convert PDF to images
    images = convert_from_path(pdf_path, 500)  # Adjust the DPI as needed

    # Extract text from each image using Tesseract OCR
    extracted_text = ""
    for i, image in enumerate(images):
        image_path = f"temp_image_{i}.png"
        image.save(image_path, 'PNG')

        reduced_noise_image_path = f"reduced_noise_temp_image_{i}.png"
        reduced_noise_img = reduce_noise(image_path)
        reduced_noise_img.save(reduced_noise_image_path, 'PNG')
        
        text = pytesseract.image_to_string(image, "vie")
        extracted_text += f"Page {i + 1} text:\n{text}\n\n"

    return extracted_text

def reduce_noise(image_path):
    # Read the image
    image = cv2.imread(image_path)

    # Convert the image to grayscale
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Apply Gaussian blur to the grayscale image
    blurred_image = cv2.GaussianBlur(gray_image, (5, 5), 0)

    # Apply adaptive thresholding to the blurred image
    thresholded_image = cv2.adaptiveThreshold(blurred_image, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 11, 2)

    # Display the original and processed images
    # cv2.imshow('Original Image', image)
    # cv2.imshow('Reduced Noise Image', thresholded_image)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()

    return thresholded_image

# Replace "path/to/your/image.jpg" with the actual path to your image
# for i in range(0, 11):
    # image_path = f"temp_image_{i}.png"
    # reduce_noise(image_path)

# image_path = "temp_image_11.png"
# reduce_noise(image_path)
# text = pytesseract.image_to_string(reduce_noise_image)
# print(text)

# Replace "path/to/your/scanned.pdf" with the actual path to your scanned PDF
pdf_path = "scanned_tax_codes.pdf"
text_result = extract_text_from_scanned_pdf(pdf_path)
print(text_result)


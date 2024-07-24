import pytesseract
from pdf2image import convert_from_path
from docx import Document
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed
import tqdm  # For progress bar

# Configure Tesseract executable path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

pdf_path = 'geographical150.pdf'
poppler_path = r"D:\Ministry of Rural Development Intern\package\poppler-24.07.0\Library\bin"
batch_size = 10  # Number of pages to process in each batch

# Custom Tesseract configuration
custom_config = r'-l khm+eng --oem 3 --psm 6'

def process_image(image, i):
    """Process image and perform OCR."""
    text = pytesseract.image_to_string(image, config=custom_config)
    return f'Extracted text from page {i+1}:\n{text}'

def process_batch(start_page, end_page):
    """Process a batch of pages."""
    batch_images = convert_from_path(pdf_path, poppler_path=poppler_path, first_page=start_page + 1, last_page=end_page)
    results = []
    for i, image in enumerate(batch_images, start=start_page):
        try:
            result = process_image(image, i)
            results.append(result)
        except Exception as e:
            results.append(f'Error processing page {i+1}: {e}')
    return results

# Initialize a Document object to store the extracted text
doc = Document()

# Get the total number of pages
images = convert_from_path(pdf_path, poppler_path=poppler_path, dpi=150)  # Initial conversion to get page count
total_pages = len(images)

# Use ThreadPoolExecutor to process pages in batches
with ThreadPoolExecutor() as executor:
    futures = []
    with tqdm.tqdm(total=total_pages, desc="Processing Pages", unit="page") as pbar:
        # Submit batches for processing
        for start_page in range(0, total_pages, batch_size):
            end_page = min(start_page + batch_size - 1, total_pages - 1)
            futures.append(executor.submit(process_batch, start_page, end_page))

        # Process results and update progress
        for future in as_completed(futures):
            batch_results = future.result()
            for result in batch_results:
                doc.add_paragraph(result)
            pbar.update(len(batch_results))

# Save the DOC file with the extracted text from each page
doc.save('extracted_text.docx')

print("Text extraction complete!")

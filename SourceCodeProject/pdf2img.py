import os
from pdf2image import convert_from_path

poppler_path = r'H:/OCR/Popler/poppler-23.07.0/Library/bin'
pdf_folder = r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\TEST_INVOICE_PDF"
saving_folder = r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Compare_Invoice"
os.makedirs(saving_folder, exist_ok=True)
for pdf_filename in os.listdir(pdf_folder):
    if pdf_filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, pdf_filename)
        pages = convert_from_path(pdf_path=pdf_path, poppler_path=poppler_path)
        for c, page in enumerate(pages, start=1):
            img_name = f"{os.path.splitext(pdf_filename)[0]}_Page{c}.png"
            img_path = os.path.join(saving_folder, img_name)
            page.save(img_path, "png")
    os.remove(pdf_path)
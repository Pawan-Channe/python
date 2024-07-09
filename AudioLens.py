
import PyPDF2
import pyttsx3
import pytesseract
from PIL import Image

def extract_text_from_image(image_path):
    image = Image.open(image_path)
    imgtxt = pytesseract.image_to_string(image)
    return imgtxt

def extract_text_from_pdf(pdf_path):
    pdftxt = ""
    with open(pdf_path, 'rb') as path:
        pdfReader = PyPDF2.PdfReader(path)
        for page_num in range(len(pdfReader.pages)):
            page = pdfReader.pages[page_num]
            pdftxt += page.extract_text()
    return pdftxt

def read_text_aloud(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

# Change the path you want to read
pathtoread = r"C:\Users\pawan\Downloads\some-sunshine.pdf"

if pathtoread.endswith('.pdf'):
    text = extract_text_from_pdf(pathtoread)
else:
    try:
        text = extract_text_from_image(pathtoread)
    except:
        print('Unsupported Format')

print(text)
read_text_aloud(text=text)

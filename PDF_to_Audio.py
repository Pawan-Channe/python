
import PyPDF2
import pyttsx3

path = open(r'C:\Users\pchanne1\Downloads\Give Me Some Sunshine Lyrics PDF Download.pdf', 'rb')

pdfReader = PyPDF2.PdfReader(path)

from_page = pdfReader.pages[1]

text = from_page.extract_text()

engine = pyttsx3.init()
engine.say(text)

engine.runAndWait()

text_to_speak = """Hello 
Anil sir, how are you?
will do, chicken party on saturday?
"""

engine = pyttsx3.init()
engine.say(text_to_speak)
engine.runAndWait()


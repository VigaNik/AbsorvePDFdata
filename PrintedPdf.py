import sys
import fitz
import xml.etree.ElementTree as ET
import os

def extraxtPDFtext(pdf_text):
    doc = fitz.open(pdf_text)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text("text")

    return text



if __naim__ == "__main__":

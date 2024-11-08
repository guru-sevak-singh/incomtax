import fitz  # PyMuPDF
import os



def read_pdf_15CA(file_path):
    pdf_document = fitz.open(file_path)

    page = pdf_document[0]
    text = page.get_text()
    text = text.split("Acknowledgement Number")[1]
    text = text.split("Part")[0]
    text = text.replace("-", "")
    text = text.replace(" ", "")
    text = text.replace("\n", "")
    
    pdf_document.close()
    
    return text

def read_pdf_15CB(file_path):
    pdf_document = fitz.open(file_path)

    page = pdf_document[0]
    text = page.get_text()
    text = text.split("Acknowledgement Number")[1]
    text = text.split("We")[0]
    text = text.replace("-", "")
    text = text.replace(" ", "")
    text = text.replace("\n", "")
    
    pdf_document.close()
    
    return text


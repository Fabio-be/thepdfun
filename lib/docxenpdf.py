import os
from docx2pdf import convert

def convertir_word_pdf(input_path,output_path):
    convert(input_path,output_path)

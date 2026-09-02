from pdf2docx import Converter

def pdf_to_docx(input_file,output_file):
        conv=Converter(pdf_file=input_file)
        conv.convert(output_file)
        conv.close()


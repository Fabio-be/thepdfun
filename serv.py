from flask import Flask
from flask import render_template
from flask import request
from flask import send_file
from werkzeug.utils import secure_filename
from lib.pdfendocx import pdf_to_docx
from lib.docxenpdf import convertir_word_pdf
from lib.pptxtopdf import pptx_to_pdf
from lib.imgtopdf import convert_img
from lib.pdfenppt import convert_toppt
from lib.delete import delete_old_files
import asyncio
import os
import threading
import time 


upload_folder="uploads"
output_folder="output"
allowed_extensions={"pdf","docx","pptx","jpg","png",".xlsx"}

app=Flask(__name__)

app.config["UPLOAD_FOLDER"]=upload_folder
app.config["OUTPUT_FOLDER"]=output_folder

os.makedirs(upload_folder,exist_ok=True)
os.makedirs(output_folder,exist_ok=True)

def allowed_file(filename):
    return "." in filename and filename.rsplit(".",1)[1].lower() in allowed_extensions

delete_old_files("/output")
delete_old_files("/uploads")

@app.route("/")
def home():
    return render_template("home.html")

@app.errorhandler(404)
def error_handler(error):
    return render_template("404.html"),404

# pdf to word
@app.route("/pdftoword",methods=["GET"])
def pdf_to_word():
    return render_template("pdftoword.html")

@app.route("/pdftoword",methods=["POST"])
def form():
    if "file" not in request.files:
        return "erreur",404
    file=request.files["file"]
    
    if file.filename=="":
        return render_template("pdftoword.html",error="fichier non trouvé")
    
    if file and allowed_file(file.filename):
        filename=secure_filename(file.filename)
        pdfpath=os.path.join(app.config["UPLOAD_FOLDER"],filename)
        file.save(pdfpath)

        docx=filename.rsplit('.', 1)[0] + '.docx'
        docx_path=os.path.join(app.config["OUTPUT_FOLDER"],docx)
        pdf_to_docx(pdfpath,docx_path)
        return send_file(docx_path,as_attachment=True)

# word to pdf ok

@app.route("/wordtopdf",methods=["GET"])
def word_topdf():
    return render_template("wordtopdf.html")

@app.route("/wordtopdf",methods=["POST"])
def word_topdf_form():
    if "file" not in request.files:
        return render_template("wordtopdf.html",error="File not found")
    file=request.files["file"]
    if file.filename=="":
        return render_template("wordtopdf.html",error="Please select file")
    
    if file and allowed_file(file.filename):
        filename=secure_filename(file.filename)
        docxpath=os.path.join(app.config["UPLOAD_FOLDER"],filename)
        file.save(docxpath)

        pdf=filename.rsplit('.',1)[0]+'.pdf'
        pdfpath=os.path.join(app.config["OUTPUT_FOLDER"],pdf)
        convertir_word_pdf(docxpath,pdfpath)
        return send_file(pdfpath,as_attachment=True)
    
# pptx to pdf
@app.route("/pptxtopdf",methods=["GET"])
def pptx():
    return render_template("pptxtopdf.html")

@app.route("/pptxtopdf",methods=["POST"])
def convert_pptx():
    if "file" not in request.files:
        return render_template("pptxtopdf.html",erreur="No file")
    file=request.files["file"]
    if file.filename=="":
        return render_template("pptxtopdf.html",error="No file")
    
    if file and allowed_file(file.filename):
        filename=secure_filename(file.filename)
        pptxpath=os.path.join(app.config["UPLOAD_FOLDER"],filename)
        file.save(pptxpath)

        pdf=filename.rsplit('.',1)[0]+'.pdf'
        pdfpath=os.path.join(app.config["OUTPUT_FOLDER"],pdf)
        pptx_to_pdf(pptxpath,pdfpath)
        return send_file(pdfpath,as_attachment=True)

# img to pdf ok
@app.route("/imgtopdf",methods=["GET"])
def img_pdf():
    return render_template("imgtopdf.html")

@app.route("/imgtopdf",methods=["POST"])
def img_topdf():
    if "file" not in request.files:
        return render_template("imgtopdf.html",error="No file")
    file=request.files["file"]
    if file.filename=="":
        return render_template("imgtopdf.html",error="No file")
    if file and allowed_file(file.filename):
        filename=secure_filename(file.filename)
        filepath=os.path.join(app.config["UPLOAD_FOLDER"],filename)
        file.save(filepath)

        pdf=filename.rsplit(".",1)[0]+".pdf"
        pdfpath=os.path.join(app.config["OUTPUT_FOLDER"],pdf)
        convert_img(filepath,pdfpath)
        return send_file(pdfpath,as_attachment=True)

# pdf en ppt ok
@app.route("/pdftopptx",methods=["GET"])
def page_ppt():
    return render_template("pdftopptx.html")

@app.route("/pdftopptx",methods=["POST"])
def form_ppt():
    if "file" not in request.files:
        return render_template("pdftopptx.html",error="No file")
    file=request.files["file"]

    if file.filename=="":
        return render_template("pdftopptx.html",error="No file")
    if file and allowed_file(file.filename):
        filename=secure_filename(file.filename)
        filepath=os.path.join(app.config["UPLOAD_FOLDER"],filename)
        file.save(filepath)

        pptname=filename.rsplit(".",1)[0]+".pptx"
        pptpath=os.path.join(app.config["OUTPUT_FOLDER"],pptname)
        convert_toppt(filepath,pptpath)
        return send_file(pptpath,as_attachment=True)
        
    
    
@app.route("/don")
def page_don():
    return render_template("don.html")
    


if __name__=="__main__":
    app.run(debug=True)
    
    

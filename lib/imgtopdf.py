from PIL import Image
import img2pdf


def convert_img(file,pdf):
    with open(pdf,"wb") as f:
        f.write(img2pdf.convert(file))


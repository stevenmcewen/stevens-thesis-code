import pytesseract
from PIL import Image


class Ocr:
    def __init__(self, image_path):
        self.image_path = image_path
        self.text = None

    def extract_text(self):
        image = Image.open(self.image_path)
        self.text = pytesseract.image_to_string(image)
        return self.text

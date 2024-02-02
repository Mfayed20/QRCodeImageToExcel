import os
from PIL import Image
from pyzbar.pyzbar import decode
from datetime import date
from multiprocessing import Pool


class QRCodeExtractor:
    def __init__(self, image_dir):
        self.image_dir = image_dir

    def extract_url(self, file_path):
        qr_image = Image.open(file_path)
        decoded_objects = decode(qr_image)
        for obj in decoded_objects:
            if obj.type == "QRCODE":
                return obj.data.decode("utf-8")
        return None

    def process_image(self, filename):
        file_path = os.path.join(self.image_dir, filename)
        url = self.extract_url(file_path)
        image_name = filename.rsplit('.', 1)[0]
        if url:
            return {
                "Image Name": image_name,
                "URL": url,
                "Date": date.today(),
                "Status": "",
            }
        else:
            return {
                "Image Name": image_name,
                "URL": None,
                "Date": date.today(),
                "Status": "URL extraction failed",
            }
 
    def process_all_images(self):
        image_files = [f for f in os.listdir(self.image_dir) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
        
        with Pool() as p:
            results = p.map(self.process_image, image_files)
        
        successful_extractions = [result for result in results if result["URL"] is not None]
        failed_extractions = [result for result in results if result["URL"] is None]
        return successful_extractions, failed_extractions

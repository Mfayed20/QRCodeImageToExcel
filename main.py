from qrcode_extractor import QRCodeExtractor
from excel_workbook_creator import ExcelWorkbookCreator
import time


def main():
    start_time = time.time()

    # Define the directory with images and the output Excel file
    image_dir = 'your_image_directory'
    output_excel = 'your_output_file.xlsx'

    # Initialize the QRCodeExtractor with the image directory
    extractor = QRCodeExtractor(image_dir)
    successful_extractions, failed_extractions = extractor.process_all_images()

    # Initialize the ExcelWorkbookCreator with the extraction results and output file
    workbook_creator = ExcelWorkbookCreator(successful_extractions, failed_extractions, output_excel)
    workbook_creator.create_workbook()

    end_time = time.time()
    execution_time = end_time - start_time
    print(f'Execution time: {execution_time:.2f} seconds')


if __name__ == '__main__':
    main()

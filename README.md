# QR Code Extractor and Excel Workbook Creator

This program extracts QR codes from images in a specified directory and outputs the results to an Excel workbook.

## How it Works

The program follows these steps:

1. Initialize the program with the directory of images and the name of the output Excel file.
2. Process all images in the directory with the `QRCodeExtractor` class.
3. Divide the results into successful and failed extractions.
4. Initialize the `ExcelWorkbookCreator` class with the successful and failed extractions and the name of the output Excel file.
5. Create an Excel workbook with two sheets: "Successful Extractions" and "Failed Extractions".
6. Populate the sheets with the respective extraction results.
7. Apply styles and formatting to the workbook.
8. Save the workbook to the output Excel file.

```mermaid
sequenceDiagram
    participant Main as "main.py"
    participant QRCodeExtractor as "QRCodeExtractor class"
    participant ExcelWorkbookCreator as "ExcelWorkbookCreator class"
    Main->>Main: Initialize start_time
    Main->>Main: Define image_dir, output_excel
    Main->>QRCodeExtractor: Initialize with image_dir
    QRCodeExtractor->>QRCodeExtractor: process_all_images()
    QRCodeExtractor-->>Main: Return successful_extractions, failed_extractions
    Main->>ExcelWorkbookCreator: Initialize with successful_extractions, failed_extractions, output_excel
    ExcelWorkbookCreator->>ExcelWorkbookCreator: create_workbook()
    Main->>Main: Calculate execution time
    Main->>Main: Print execution time
```

## Sheets Design

Successful Extractions:

```mermaid
graph TD
    A[Header Row] --> B[Image Name]
    A --> C[URL Column]
    A --> D[Date Column]
    A --> E[Status Column]
    B --> F[Image Format: .jpg, .png, etc.]
    C --> G[Hyperlinks]
    D --> H[Date Format: DD/MM/YYYY]
    E --> I[Dropdown Menu]
    I -.-> J[Options: 'Applied', 'Approved', 'Rejected', 'Not Applicable']
    style A fill:#4472C4, color:#FFFFFF
    style B fill:#f9f9f9
    style C fill:#f9f9f9
    style D fill:#f9f9f9
    style E fill:#f9f9f9
    style F fill:#f9f9f9
    style G fill:#f9f9f9
    style H fill:#f9f9f9
    style I fill:#f9f9f9
    style J fill:#f9f9f9

```

Failed Extractions:

```mermaid
graph TD
    A[Header Row] --> B[Image Name]
    B --> C[Image Format: .jpg, .png, etc.]
    style A fill:#4472C4, color:#FFFFFF
    style B fill:#f9f9f9
    style C fill:#f9f9f9
```

## Requirements

* Python 3.6 or higher
* openpyxl library for Excel workbook creation and manipulation
* opencv-python and pyzbar libraries for QR code extraction

```bash
pip install openpyxl opencv-python pyzbar Pillow
```

## Usage

To use this program, run the `main.py` script with Python 3.6 or higher.

```bash
python main.py
```

## Acknowledgements

This code was created with the assistance of OpenAI's GPT-4 model.

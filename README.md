Email Attachment Processor

Description-
This Python project automates the process of reading email attachments, extracting key details such as receipt number, vendor name, date, and total, and then sends a notification email to the finance team summarizing the extracted details.

Features-
Extracts key details from email attachments automatically.
Supports receipt data such as: Receipt Number, Vendor Name, Date, Total Amount
Sends a summary email to the finance team.
Improves efficiency in processing financial documents.

Prerequisites-
Python 3.6.x-3.9.x
Email credentials (for accessing and sending emails)
Required libraries (listed in requirements.txt)

How It Works-
Accesses the email inbox using IMAP.
Downloads and processes attachments (e.g., PDFs or images).
Extracts key information using libraries like PyPDF2 and PaddleOCR.
Saves the extracted details in an excel file.

Known Issues-
Ensure attachments are in supported formats (e.g., PDFs, images).
OCR accuracy may vary based on attachment quality.
bash
Copy code
python test_email_processor.py
Known Issues
Ensure attachments are in supported formats (e.g., PDFs, images).
OCR accuracy may vary based on attachment quality.

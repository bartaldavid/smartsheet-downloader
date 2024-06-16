import os
from consts import *

os.makedirs(merged_route, exist_ok=True)
os.makedirs(bank_route, exist_ok=True)
os.makedirs(temp_route, exist_ok=True)
os.makedirs(final_route, exist_ok=True)

# Check if bank route contains PDF files
receipt_pdf_files = [file for file in os.listdir(bank_route) if file.endswith(".pdf")]

if len(receipt_pdf_files) == 0:
    raise Exception("Bank route does not contain any PDF files.")

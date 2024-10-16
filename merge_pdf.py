from pypdf import PdfWriter
import os
from consts import temp_route, merged_route

cwd = os.getcwd()
os.makedirs(merged_route, exist_ok=True)

# Step 1: List all PDF files in the directory
receipt_pdf_files = [file for file in os.listdir(temp_route) if file.endswith(".pdf")]

# Step 2: Group files by the prefix before the underscore
pdf_groups: dict[str, list[str]] = {}

for file in receipt_pdf_files:
    prefix = file.split("__")[0]

    if prefix in pdf_groups:
        pdf_groups[prefix].append(file)
    else:
        pdf_groups[prefix] = [file]

# Step 3: Merge PDFs for each group
for prefix, files in pdf_groups.items():
    merger = PdfWriter()

    for file in sorted(files):
        merger.append(os.path.join(cwd, temp_route, file))

    output_filename = os.path.join(cwd, merged_route, f"{prefix}.pdf")

    merger.write(output_filename)
    merger.close()

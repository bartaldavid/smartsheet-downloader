from fnmatch import fnmatch
import shutil
from urllib.request import urlretrieve
import smartsheet.models
import logging
import os
from consts import *


class Attachment:
    def __init__(
        self,
        final_name: str,
        original_name: str,
        attachment: smartsheet.models.Attachment,
        download_url: str = None,
    ):
        self.final_name = final_name
        self.original_name = original_name
        self.attachment = attachment
        self.download_url = download_url


column_map = {}
cwd = os.getcwd()


# Helper function to find cell in a row
def get_cell_by_column_name(
    row: smartsheet.models.Row, column_name: str
) -> smartsheet.models.Cell:
    column_id = column_map[column_name]
    return row.get_column(column_id)


smart = smartsheet.Smartsheet()
# Make sure we don't miss any error
smart.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename="rwsheet.log", level=logging.INFO)

# Load entire sheet
sheet: smartsheet.models.sheet.Sheet = smart.Sheets.get_sheet(
    sheet_id=sheet_id, include="attachments"
)

# Build column map for later reference - translates column names to column id
for column in sheet.columns:
    column_map[column.title] = column.id


rows: list[smartsheet.models.row.Row] = sheet.rows
files: list[Attachment] = []
receipt_pdf_files = [file for file in os.listdir(bank_route) if file.endswith(".pdf")]


for row in rows:
    # print(row)
    op_number = get_cell_by_column_name(row, "OP elsz.sorszám").display_value
    ready_state = get_cell_by_column_name(row, "Ready")
    primary = get_cell_by_column_name(row, "Primary Column")
    receipt_number = get_cell_by_column_name(row, "Kifizetés bizonylata").display_value

    if ready_state.value != "Green":
        continue

    if receipt_number is not None and not receipt_number.startswith("P"):
        final_receipt_filename = f"{op_number}_b_{receipt_number}.pdf"
        found = False

        for receipt_file in receipt_pdf_files:
            if fnmatch(receipt_file, f"{receipt_number}_*.pdf"):
                # shutil.copy(f'{bank_route}/{receipt_file}', f'{temp_route}/{final_receipt_filename}')
                shutil.copy(
                    os.path.join(cwd, bank_route, receipt_file),
                    os.path.join(cwd, temp_route, final_receipt_filename),
                )
                found = True
                break

        print(f'Receipt "{receipt_number}" not found for OP {op_number}')

    for att in row.attachments:
        final_name = f"{op_number}__{att.name}"
        files.append(Attachment(final_name, att.name, att, download_url=None))



for file in files:
    data: smartsheet.models.Attachment = smart.Attachments.get_attachment(
        sheet_id=sheet_id, attachment_id=file.attachment.id
    )
    file.download_url = data.url
    urlretrieve(data.url, os.path.join(cwd, temp_route, file.final_name))
import pandas as pd
import os
import shutil
from consts import *

cwd = os.getcwd()

df = pd.read_excel(os.path.join(cwd, xlsx_name))

final_files = os.listdir(os.path.join(cwd, merged_route))

for file in final_files:
    prefix = int(file.split("_")[0])
    if prefix not in df["Document No."].values:
        print(f"OP {prefix} not found in the Excel file.")
        continue
    row = df.loc[df["Document No."] == prefix].iloc[0]
    suffix = "" if str(row["Filename"]) == "nan" else row["Filename"]
    final_filename = (
        f"{row['Country Code']}{str(row['Document No.']).zfill(3)}{suffix}.pdf"
    )
    print(final_filename)

    shutil.copy(
        os.path.join(cwd, merged_route, file),
        os.path.join(cwd, final_route, final_filename),
    )

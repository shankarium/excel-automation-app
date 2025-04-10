import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import io

def process_excel(df):
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    red_font = Font(color="FF0000")

    for row in ws.iter_rows():
        for cell in row:
            cell.font = red_font

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

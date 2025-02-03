# mergeExcelFiles.py
# Created by David Herren
# Version 2024-12-09

import pandas as pd
import os
from tkinter import filedialog, Tk
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

def auto_adjust_column_widths(ws, df):
  ws.column_dimensions['A'].width = 30
  for col_idx, col_name in enumerate(df.columns, start=1):
    if col_idx == 1:
      continue
    max_length = max((len(str(cell)) for cell in df[col_name]), default=0)
    ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 2

def mergeExcelFiles():
  root = Tk()
  root.withdraw()

  # File selection
  file_paths = filedialog.askopenfilenames(title='Bitte wählen Sie die Excel-Dateien aus, die zusammengeführt werden sollen')
  if not file_paths:
    print("Keine Dateien ausgewählt. Skript wird beendet.")
    return

  # Read and preprocess files
  all_data = []
  filenames = []
  expected_columns = ['Bezeich.-Désignation', 'Mat.', 'Bem.-Remarques', 'Stk', 'Art.No.']

  for file_path in file_paths:
    norm_path = os.path.normpath(file_path)
    df = pd.read_excel(norm_path)
    filenames.append(os.path.basename(norm_path))

    # Add missing columns at once
    for col in expected_columns:
      if col not in df.columns:
        df[col] = None

    if 'Pos.' in df.columns:
      df = df.drop(columns=['Pos.'])

    all_data.append(df)

  combined_df = pd.concat(all_data, ignore_index=True)

  # Group data by Art.No. and Bem.-Remarques
  grouped = combined_df.groupby(['Art.No.', 'Bem.-Remarques'], dropna=False).agg({
    'Stk': 'sum',
    'Bezeich.-Désignation': 'first',
    'Mat.': 'first'
  })

  count_df = combined_df.groupby(['Art.No.', 'Bem.-Remarques'], dropna=False).size().rename('Count')
  grouped = grouped.merge(count_df, left_index=True, right_index=True).reset_index()

  # Set column order and mark combined lines
  columns_order = ['Art.No.', 'Stk', 'Bezeich.-Désignation', 'Mat.', 'Bem.-Remarques', 'Count']
  grouped = grouped[columns_order]
  grouped[''] = grouped['Count'].apply(lambda x: '*' if x > 1 else '')
  grouped = grouped.drop(columns='Count')

  final_order = ['Art.No.', 'Stk', 'Bezeich.-Désignation', 'Mat.', 'Bem.-Remarques', '']
  grouped = grouped[final_order]

  # Ask for filename to save
  new_file_name = filedialog.asksaveasfilename(title='Bitte geben Sie den Namen für das neue Excel-File an', filetypes=[('Excel Files', '*.xlsx')])
  if not new_file_name:
    print("Kein Dateiname eingegeben. Skript wird beendet.")
    return

  # Write to Excel
  wb = Workbook()
  ws = wb.active

  header_description = 'Zusammenfassung der Stk-Listen von ' + ', '.join(filenames)
  ws.append([header_description])

  for r in dataframe_to_rows(grouped, index=False, header=True):
    ws.append(r)

  # Adjust column widths
  auto_adjust_column_widths(ws, grouped)

  wb.save(new_file_name + '.xlsx')
  print(f"Datei wurde erfolgreich gespeichert: {new_file_name}.xlsx")

if __name__ == "__main__":
  mergeExcelFiles()

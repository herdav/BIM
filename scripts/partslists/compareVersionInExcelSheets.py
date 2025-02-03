# compareVersionInExcelSheets.py
# Created by David Herren
# Version 2024-06-18

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from openpyxl.styles import Font, Border

def select_file(title):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])
    return file_path

def auto_adjust_column_width(worksheet, columns):
    for col in columns:
        max_length = 0
        column = worksheet[col]
        for cell in column:
            if cell.coordinate in worksheet.merged_cells:
                continue
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[col].width = adjusted_width

def main():
    try:
        old_file = select_file("Select the old Excel file")
        new_file = select_file("Select the new Excel file")

        old_file_name = Path(old_file).name
        new_file_name = Path(new_file).name

        # Spaltennamen in der zweiten Zeile -> header=1
        df_old = pd.read_excel(old_file, header=1)
        df_new = pd.read_excel(new_file, header=1)

        # Spalten bereinigen
        df_old.columns = df_old.columns.str.strip()
        df_new.columns = df_new.columns.str.strip()

        required_columns = ['Art.No.', 'Bem.-Remarques', 'Stk', 'Bezeich.-Désignation']
        for df in [df_old, df_new]:
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Missing columns in data: {missing_columns}")

        # Leerzeichen entfernen und NaN bzw. 'nan' durch leere Strings ersetzen
        df_old['Art.No.'] = df_old['Art.No.'].astype(str).str.strip().replace('nan', '')
        df_old['Bem.-Remarques'] = df_old['Bem.-Remarques'].astype(str).str.strip().replace('nan', '')
        df_new['Art.No.'] = df_new['Art.No.'].astype(str).str.strip().replace('nan', '')
        df_new['Bem.-Remarques'] = df_new['Bem.-Remarques'].astype(str).str.strip().replace('nan', '')

        # Pseudoartikel erstellen
        df_old['Pseudoartikel'] = df_old['Art.No.'].astype(str) + '|' + df_old['Bem.-Remarques'].astype(str)
        df_new['Pseudoartikel'] = df_new['Art.No.'].astype(str) + '|' + df_new['Bem.-Remarques'].astype(str)

        # Spalten filtern
        df_old = df_old[['Pseudoartikel', 'Stk', 'Bezeich.-Désignation']]
        df_new = df_new[['Pseudoartikel', 'Stk', 'Bezeich.-Désignation']]

        # Merge
        df_merged = pd.merge(df_old, df_new, on='Pseudoartikel', how='outer', suffixes=('_old', '_new'), indicator=True)

        # Änderungen kennzeichnen
        df_merged['Änderung'] = df_merged.apply(
            lambda row: 'Entfernt' if row['_merge'] == 'left_only' else
                        'Hinzugefügt' if row['_merge'] == 'right_only' else
                        ('Geändert' if row['Stk_old'] != row['Stk_new'] else 'Keine Änderung'),
            axis=1
        )

        # Pseudoartikel wieder aufteilen
        df_merged[['Art.No.', 'Bem.-Remarques']] = df_merged['Pseudoartikel'].str.split('|', expand=True)
        df_merged = df_merged.drop(columns=['Pseudoartikel', '_merge'])

        # Bemerkungen nochmals bereinigen, falls noch 'nan' vorhanden sein sollte
        df_merged['Bem.-Remarques'] = df_merged['Bem.-Remarques'].replace('nan', '')

        # Bezeichnung bei Entfernt aktualisieren
        df_merged['Bezeich.-Désignation'] = df_merged.apply(
            lambda row: row['Bezeich.-Désignation_old'] if row['Änderung'] == 'Entfernt' else row['Bezeich.-Désignation_new'],
            axis=1
        )

        # Sortieren (Hinzugefügt -> Geändert -> Entfernt -> Keine Änderung)
        sort_order = {'Hinzugefügt': 0, 'Geändert': 1, 'Entfernt': 2, 'Keine Änderung': 3}
        df_merged = df_merged.sort_values(
            by=['Änderung', 'Art.No.', 'Bem.-Remarques'],
            key=lambda col: col.map(sort_order).fillna(col)
        )

        # Spaltennamen anpassen: Stk_old -> Stk.-alt, Stk_new -> Stk.-neu
        df_merged.rename(columns={'Stk_old': 'Stk.-alt', 'Stk_new': 'Stk.-neu'}, inplace=True)

        # Gewünschte Spaltenauswahl
        df_merged = df_merged[['Änderung', 'Stk.-alt', 'Stk.-neu', 'Art.No.', 'Bezeich.-Désignation', 'Bem.-Remarques']]

        with pd.ExcelWriter(new_file, engine='openpyxl', mode='a') as writer:
            if 'Änderungen' in writer.book.sheetnames:
                del writer.book['Änderungen']

            df_merged.to_excel(writer, sheet_name='Änderungen', index=False)
            ws = writer.sheets['Änderungen']

            # Erste Zeile einfügen
            ws.insert_rows(1)
            ws.cell(row=1, column=1, value='Verglichen:')
            ws.cell(row=1, column=2, value=f'{old_file_name} vs {new_file_name}')

            # Spaltenbreiten anpassen
            auto_adjust_column_width(ws, ['A', 'D', 'E', 'F'])

            # Kopfzeile (Zeile 2) und Vergleiche-Zeile (Zeile 1) neutral formatieren
            for row in ws.iter_rows(min_row=1, max_row=2):
                for cell in row:
                    cell.font = Font(bold=False, color='000000')
                    cell.border = Border()

            # Datenzeilen formatieren (ab Zeile 3)
            for row in range(3, ws.max_row + 1):
                change_type = ws.cell(row=row, column=1).value
                if change_type == 'Hinzugefügt':
                    font_color = 'FF0000'  # Rot
                elif change_type == 'Geändert':
                    font_color = '00B0F0'  # Blau
                elif change_type == 'Entfernt':
                    font_color = 'FFC000'  # Gelb
                else:
                    font_color = '000000'  # Schwarz für Keine Änderung

                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.font = Font(color=font_color, bold=False)
                    cell.border = Border()

            writer.book.save(new_file)

        messagebox.showinfo("Erfolg", "Vergleich erfolgreich abgeschlossen und die Datei wurde gespeichert.")

    except Exception as e:
        messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

if __name__ == "__main__":
    main()

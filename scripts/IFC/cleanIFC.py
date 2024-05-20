# cleanIFC.py
# Version 20.05.2024 by David Herren @ WiVi AG

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def replace_strings_in_file(filename, is_abbruch):
  if not os.path.exists(filename):
    print(f"Die Datei {filename} wurde nicht gefunden.")
    return None, 0, 0

  with open(filename, 'r', encoding='utf-8') as file:
    content = file.read()

  if is_abbruch:
    content = content.replace("0.30196078431372547,0.30196078431372547,0.30196078431372547", "1.0,0.8431372549019608,0.0")

  default_replacement = None
  pty_replacement_count = 0  # Zählt die Anzahl der Ersetzungen für PTY_
  wiv_replacement_count = 0  # Zählt die Anzahl der Ersetzungen für WIV_

  def find_and_replace(pattern):
    matches = re.finditer(pattern, content)
    replacements = []
    for match in matches:
      full_match = match.group(0)
      value_match = match.group(2)
      try:
        value = float(value_match)
        rounded_value = round(value * 30.48, 3)
        replacement = f"{match.group(1)}{rounded_value}{match.group(3)}"
        replacements.append((full_match, replacement))
      except ValueError:
        pass
    return replacements

  # Ersetze Kilometrierung und Epsilon
  patterns = {
    "Kilometrierung": r"(IFCPROPERTYSINGLEVALUE\('Kilometrierung',[^,]+,IFCREAL\()([-0-9\.]+)(\))",
    "Epsilon": r"(Epsilon',\s*\$,IFCREAL\()([-0-9\.]+)(\))"
  }

  for key, pattern in patterns.items():
    replacements = find_and_replace(pattern)
    for old, new in replacements:
      content = content.replace(old, new)

  # Ersetze PTY_[int] und WIV_[int]
  for prefix in ['PTY', 'WIV']:
    pattern = fr'{prefix}_\d+ '
    count = len(re.findall(pattern, content))
    content = re.sub(pattern, '', content)
    if prefix == 'PTY':
      pty_replacement_count += count
    else:
      wiv_replacement_count += count

  # Überprüfe und ersetze "Default"
  if "Default" in content:
    replace_default = messagebox.askyesno("Ersetzen", "Soll die Zeichenkette 'Default' ersetzt werden?")
    if replace_default:
      replacement = simpledialog.askstring("Ersetzen", "Geben Sie die Ersatzeichenkette für 'Default' ein:")
      if replacement is None or not replacement.strip():
        return None, 0, 0
      content = content.replace("Default", replacement)
      default_replacement = f"Default → {replacement}"

  # Benutzerdefinierte Zeichenfolgenersetzung
  custom_replace = messagebox.askyesno("Benutzerdefinierte Ersetzung", "Möchten Sie eine benutzerdefinierte Zeichenfolge ersetzen?")
  if custom_replace:
    target_string = simpledialog.askstring("Zielzeichenfolge", "Geben Sie die zu ersetzende Zeichenfolge ein:")
    if target_string:
      replacement_string = simpledialog.askstring("Ersatzeichenfolge", "Geben Sie die Ersatzeichenfolge ein:")
      if replacement_string is not None:  # Prüfen, ob der Benutzer nicht auf "Abbrechen" geklickt hat.
        content = content.replace(target_string, replacement_string)
      else:
        # Der Benutzer hat auf "Abbrechen" geklickt oder keine Eingabe getätigt.
        messagebox.showinfo("Abbruch", "Benutzerdefinierte Ersetzung wurde nicht durchgeführt.")
    else:
      # Der Benutzer hat auf "Abbrechen" geklickt oder keine Eingabe getätigt.
      messagebox.showinfo("Abbruch", "Benutzerdefinierte Ersetzung wurde nicht durchgeführt.")

  confirm_changes = messagebox.askyesno("Bestätigen", "Möchten Sie die Änderungen wirklich vornehmen?")
  if not confirm_changes:
    return None, 0, 0

  return content, pty_replacement_count, wiv_replacement_count

if __name__ == "__main__":
  root = tk.Tk()
  root.withdraw()

  # Dialogfenster für die Dateiauswahl
  filepaths = filedialog.askopenfilenames(title="Wählen Sie die zu bereinigenden Dateien aus",
                                          filetypes=[("IFC Dateien", "*.ifc")])
  if not filepaths:
    exit()

  # Fenster zur Auswahl des Abbruchmodells
  abbruch_model = messagebox.askyesno("Abbruchmodell", "Handelt es sich um ein Abbruchmodell?")

  default_replacement = None
  total_pty_replacement_count = 0
  total_wiv_replacement_count = 0

  for filepath in filepaths:
    content, pty_repl_count, wiv_repl_count = replace_strings_in_file(filepath, abbruch_model)
    if content is None:
      continue
    total_pty_replacement_count += pty_repl_count
    total_wiv_replacement_count += wiv_repl_count
    with open(filepath, 'w', encoding='utf-8') as file:
      file.write(content)

  output = f"Es wurden {total_pty_replacement_count} 'PTY_' und {total_wiv_replacement_count} 'WIV_' Ersetzungen vorgenommen."
  if default_replacement:
    output = f"{default_replacement}\n{output}"

  messagebox.showinfo("Bereinigungsbericht", output)

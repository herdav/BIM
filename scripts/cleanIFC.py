# Dateiname: cleanIFC.py
# Created 2023 by David Herren

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog  # Importiere zusätzliche Dialoge

def replace_strings_in_file(filename):
    # Überprüfe, ob die Datei existiert
    if not os.path.exists(filename):
        print(f"Die Datei {filename} wurde nicht gefunden.")
        return None, None, None

    # Lese den Inhalt der Datei
    with open(filename, 'r', encoding='utf-8') as file:
        content = file.read()

    default_replacement = None
    pty_replacements = []

    # Finde alle Vorkommen der Zeichenkette mit einem regulären Ausdruck
    pty_matches = re.findall(r'PTY_\d+ ', content)
    if pty_matches:
        pty_replacements.extend(pty_matches)  # Füge jede gefundene Zeichenkette der Liste hinzu
        # Ersetze die Zeichenkette mit einem regulären Ausdruck
        content = re.sub(r'PTY_\d+ ', '', content)

    # Überprüfe, ob die Zeichenkette "Default" vorkommt
    if "Default" in content:
        # Frage den Benutzer, ob "Default" ersetzt werden soll
        replace_default = messagebox.askyesno("Ersetzen", "Soll die Zeichenkette 'Default' ersetzt werden?")
        if replace_default:
            # Frage den Benutzer nach der Ersatzeichenkette
            replacement = simpledialog.askstring("Ersetzen", "Geben Sie die Ersatzeichenkette für 'Default' ein:")
            if replacement is None or not replacement.strip():  # Wenn der Benutzer den Dialog abbricht oder nichts eingibt
                return None, None, None
            content = content.replace("Default", replacement)
            default_replacement = f"Default → {replacement}\n\n"  # Zwei Zeilenumbrüche hinzugefügt

    # Frage den Benutzer, ob die Änderungen vorgenommen werden sollen
    confirm_changes = messagebox.askyesno("Bestätigen", "Möchten Sie die Änderungen wirklich vornehmen?")
    if not confirm_changes:
        return None, None, None

    # Gib den geänderten Inhalt und die Ersetzungen zurück
    return content, default_replacement, pty_replacements

if __name__ == "__main__":
    # Erstelle ein Hauptfenster (wird später versteckt)
    root = tk.Tk()
    root.withdraw()  # Verstecke das Hauptfenster

    # Öffne ein Dialogfenster zur Dateiauswahl
    filepaths = filedialog.askopenfilenames(title="Wählen Sie die zu bereinigenden Dateien aus",
                                            filetypes=[("IFC Dateien", "*.ifc")])
    if not filepaths:  # Wenn der Benutzer den Dialog abbricht
        exit()

    default_replacement = None
    all_pty_replacements = []
    # Iteriere durch die ausgewählten Dateien und bereinige sie
    for filepath in filepaths:
        content, default_repl, pty_repls = replace_strings_in_file(filepath)
        if content is None:  # Wenn der Benutzer den Dialog abgebrochen hat
            continue
        if default_repl:
            default_replacement = default_repl
        all_pty_replacements.extend(pty_repls)

        # Speichere die geänderte Datei
        with open(filepath, 'w', encoding='utf-8') as file:
            file.write(content)

    # Organisiere die Ausgabe
    output = []
    if default_replacement:
        output.append(default_replacement)
    output.extend(all_pty_replacements)

    # Zeige ein Fenster mit den bereinigten Zeichenketten, getrennt durch ein Leerzeichen
    replacement_strings = " ".join(output)
    messagebox.showinfo("Bereinigte Zeichenketten", f"Folgende Änderungen wurden vorgenommen:\n\n{replacement_strings}")

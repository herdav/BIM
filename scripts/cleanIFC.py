# Dateiname: cleanIFC.py

import os
import re  # Importiere das Modul für reguläre Ausdrücke

def replace_strings_in_file(filename):
    # Überprüfe, ob die Datei existiert
    if not os.path.exists(filename):
        print(f"Die Datei {filename} wurde nicht gefunden.")
        return

    # Lese den Inhalt der Datei
    with open(filename, 'r', encoding='utf-8') as file:
        content = file.read()

    # Ersetze die Zeichenkette mit einem regulären Ausdruck
    content = re.sub(r'PTY_\d+ ', '', content)  # Ersetze "PTY_[eine Zahl] " durch ""

    # Schreibe den geänderten Inhalt zurück in die Datei
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(content)

    print(f"'PTY_[eine Zahl] ' wurde in {filename} entfernt.")

if __name__ == "__main__":
    current_directory = os.path.dirname(os.path.abspath(__file__))

    # Iteriere durch alle Dateien im aktuellen Verzeichnis
    for filename in os.listdir(current_directory):
        if filename.endswith('.ifc'):
            filepath = os.path.join(current_directory, filename)
            replace_strings_in_file(filepath)

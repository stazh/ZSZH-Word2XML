Das Python-Skript *convert_rrb_word_to_xml.py* konvertiert Regierungsratsbeschlüsse aus Word-Dokumenten (`.docx`) in TEI-XML-Dateien und ergänzt diese basierend auf Metadaten, die aus einer Excel-Datei (`.xlsx`) geladen werden. Zusätzlich wird eine Fehlerdatei im Excel-Format erstellt, falls während der Verarbeitung Probleme auftreten.

---

## Voraussetzungen

### Python-Bibliotheken

Stelle sicher, dass folgende Python-Bibliotheken installiert sind:

- **`python-docx`**: Für das Verarbeiten von Word-Dateien.
- **`pandas`**: Für das Verarbeiten von Excel-Dateien.
- **`xlsxwriter`**: Für das Schreiben von Fehlerprotokollen.
- **`BeautifulSoup`**: Für die XML-Bearbeitung und -Erstellung.
- **`os` und `shutil`**: Für die Dateiverwaltung (bereits in Python enthalten).

Installiere alle notwendigen Pakete mit folgendem Befehl:

```bash
pip install python-docx pandas xlsxwriter beautifulsoup4
```

## Verwendung

Das Skript benötigt zwei Eingabeparameter:

1. **Input-Ordner**: Der Pfad zu einem Ordner, der die Word-Dateien mit den RRB in Unterordneren enthält, die verarbeitet werden sollen.
2. **Excel-Datei**: Der Pfad zu einer Excel-Datei, die zusätzliche die Signaturen und zugehörigen Links ins Archivinformationssystem (AIS) enthält.

### Kommandozeilenausführung

Führe das Skript über die Kommandozeile aus:

```bash
python convert_rrb_word_to_xml.py <input_folder> <metadata_file>
```

## Fehlerbehandlung

Das Skript überprüft automatisch:
- Ob die Word-Dateien korrekt verarbeitet werden können.

Falls Fehler auftreten, werden diese in `errorfile.xlsx` dokumentiert.

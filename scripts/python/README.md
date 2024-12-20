Das Python-Skript *convert_rrb_word_to_xml.py* konvertiert Word-Dokumente (`.docx`) mit Regierungsratsbeschlüssen (RRB) in TEI-XML-Dateien und ergänzt diese basierend auf Metadaten, die aus einer Excel-Datei (`.xlsx`) geladen werden. Zusätzlich wird eine Fehlerdatei im Excel-Format erstellt, falls während der Verarbeitung Probleme auftreten.

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
2. **Excel-Datei**: Der Pfad zu einer Excel-Datei, die die Signaturen (Spalte mit Name `Signatur`) und die dazugehörenden Scope IDs (Spalte mit Name `ID`) enthält.

### Kommandozeilenausführung

Führe das Skript über die Kommandozeile aus:

```bash
python convert_rrb_word_to_xml.py <input_folder> <metadata_file>
```

## Fehlerbehandlung

Das Skript überprüft automatisch, ob die Word-Dateien korrekt verarbeitet werden können. Falls Fehler auftreten, werden diese in `errorfile.xlsx` dokumentiert.


## Mögliche Weiterentwicklungen / Fehlerbehebungen

- **Ersetzen der XML-Datei, wenn diese bereits vorhanden ist**: Im Moment werden die Daten bei einer erneuten analogen Anwendung des Skripts nocheinmal in das bereits bestehende XML geschrieben. 
- **Fehlerhandling und Validierung**: Verbesserung des Errorhandlings und zusätzlich eine Validierung nach TEI-Standard durchführen.
- **Bessere Output-Ordnerstruktur**
- **Behandlung von `<p/>`-Tags**: Löschen von leeren Absätzen (`<p/>`-Tags)
- **Tabellen mit verbundenen Zellen**: Korrekte Umsetzung von Tabellen mit `colspan` und `rowspan`.
- **Zentrierte Inhalte**: Optimierung der Konvertierung von zentrierten Texten und Tabellen.
- **[up.]**: Automatische Handhabung und Konvertierung von der Seitenangabe **[up.]** (=unpaginiert) in den Metadaten.
- **Integrieren von KRP und OS**: Dieses Skript funktioniert nur für die Konvertierung von RRB. Um auch die KRP und OS damit konvertieren zu können, müssten einige Anpassungen vorgenommen werden (gewisse Metadaten sind hartkodiert).

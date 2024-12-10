### Zentrale Serien des Kantons Zürich (ZSZH)

Das Staatsarchiv des Kantons Zürich archiviert und publiziert die zentralen Dokumente der Regierungs-, Parlaments- und Verwaltungstätigkeit des Kantons Zürich seit der Gründung des modernen Staatswesens im Jahr 1803. Zu den zentralen Serien gehören:

- **Gesetzessammlung (OS)**
- **Amtsblatt (ABl)**
- **Kantonsratsprotokolle (KRP)**
- **Regierungsratsbeschlüsse (RRB)**

### Historie der Publikationslösungen

Die Transkriptionen der Serien **OS**, **KRP** und **RRB** wurden ursprünglich in Word-Templates erstellt und anschließend zusammen mit den Digitalisaten in PDF-Dateien für die Online-Publikation auf **Archives Quickaccess** bereitgestellt.

Mit der Frage nach einer zukünftigen Lösung für die Veröffentlichung des **Amtsblatts**, dessen Transkriptionen direkt im **TEI-XML**-Format erstellt werden, entschied sich das Staatsarchiv, die thematische Schnellsuche auf **Archives Quickaccess** abzulösen und alle vier Serien auf einem neuen Portal gemeinsam zu veröffentlichen (vgl. [https://www.zentraleserien.zh.ch](https://www.zentraleserien.zh.ch/home)). 

Im Zuge dieser Neugestaltung wurde beschlossen, die bisherigen Word-Dokumente durch **TEI-XML**-Dateien zu ersetzen, um eine zukunftsfähige und strukturierte Datengrundlage zu gewährleisten.

### Konvertierungsprozess

Bereits ab 2017 wurde schrittweise damit begonnen, die Word-Dokumente der Serien **OS**, **KRP** und **RRB** mittels [VBA](scripts/VBA)- und später [Python-Skripten](scripts/python) in **XML** nach dem **TEI-Standard** zu konvertieren. Diese konvertierten XML-Dateien wurden anschließend als **Open Government Data (OGD)** Datensätze veröffentlicht, die maschinenlesbar sind.

Dank der Verwendung von Word-Vorlagen konnten Metadaten wie **Titel**, **Signatur** und **Datum** automatisiert in das TEI-Format überführt werden. Zudem ermöglichte die konsistente Formatierung der Dokumente eine präzise Extraktion von Textelementen wie **Untertiteln**, **Marginalien** und **Seitenumbrüchen** mittels Mustererkennung.

### Repository-Inhalt

Dieses GitHub-Repository enthält die **Skripte** sowie **Dokumentationen** des Konvertierungsprozesses.

# 📅 Excel2ical – Schultermine als ICS exportieren

**Excel2ical** ist ein einfaches Python-Tool zur Konvertierung von Schulterminlisten aus Excel in das ICS-Format, das direkt in Outlook oder andere Kalenderprogramme importiert werden kann.
Unsere Schule benutzt einen Terminkalender, welcher als Exceldatei gepflegt wird und diese Daten sollen zusätzlich als ICS Datei für Outlook usw. exportierbar gemacht werden.

![Screenshot 2023-03-16 080438](https://user-images.githubusercontent.com/75378632/225542585-db870a1e-7f39-491b-b0e8-124f28112038.png)

Icon by Yannick (https://icon-icons.com/de/users/hao5vBiTzx3djBqoJuU6V/icon-sets/)
---

## 🚀 Funktionen

- 📤 Konvertiert Excel-Dateien (.xlsx) in ICS-Dateien (.ics), die dem Format der Beispieldatei hier im Repository entsprechen.
- 📅 Unterstützt Start- und Enddatum sowie Uhrzeiten
- 📝 Ereignisbeschreibung wird übernommen
- 🖥️ Benutzerfreundliche grafische Oberfläche (Tkinter)
- 🧠 Automatische Fehlererkennung bei ungültigen Datumswerten
- 🧩 Kompatibel mit PyInstaller für portable EXE-Erstellung

---

## 🖥️ Benutzeroberfläche

| Element                  | Beschreibung                                      |
|--------------------------|--------------------------------------------------|
| **Exceldatei auswählen** | Wählt die zu konvertierende Excel-Datei aus      |
| **ICS-Datei bestimmen**  | Legt den Namen und Speicherort der ICS-Datei fest|
| **ICS erzeugen**         | Startet die Konvertierung                        |
| **Beschreibung**         | Zeigt den Zweck des Tools an                     |
| **Logo/Icon**            | Optionales Bild zur optischen Gestaltung         |

---

## 📦 Voraussetzungen

- Python 3.x
- Module:
  - `openpyxl`
  - `icalendar`
  - `tkinter` (Standardmodul)
- Optional: `pyinstaller` zur Erstellung einer ausführbaren Datei

Installation der benötigten Pakete:

```bash
pip install openpyxl icalendar

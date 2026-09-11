import datetime
import hashlib
import os
import sys
import uuid
from openpyxl import load_workbook
from icalendar import Calendar, Event
from tkinter import (Button, Frame, Label, PhotoImage, Tk, filedialog, messagebox)

WINDOW_TITLE = "Excel2ical v2.3          (buc @ hems.de)"
WINDOW_ICON = "excel2ics.ico"
APP_TITLE = "Excel2ical"
MASCOT_PIC = "cal.png"
FONT_DESCRIPTION = ("Helvetica", 10)
FONT_BUTTONS = ("Helvetica", 12)
FONT_CONVERT_BUTTON = ("Helvetica", 12, "bold")
FONT_COLOR_CONVERT_BUTTON = "white"
FONT_LABEL = ("Helvetica", 10, "italic")
BG_COLOR = "gray80"
APP_DESCRIPTION = "Dieses Tool wandelt die Excelterminliste in eine\nOutlook-importierbare ICS-Datei um."
ICS_FILE_LABEL = "Kein ICS-Dateiname bestimmt"
ICS_BUTTON_LABEL = "ICS-Datei bestimmen"
EXCEL_FILE_LABEL = "Keine Excel-Datei ausgewählt"
EXCEL_BUTTON_LABEL = "Exceldatei auswählen"
CONVERT_BUTTON_TEXT = "ICS erzeugen"
BG_COLOR_CONVERT_BUTTON = "#ee2724"

# Punkt 1: als Modulvariablen initialisieren, damit convert_files() nicht mit
# einem NameError abstürzt, wenn "ICS erzeugen" ohne vorherige Dateiauswahl
# geklickt wird.
exc_file_path = None
ics_file_path = None

# Erwartete Mindestspaltenzahl der Terminplan-Tabelle:
# A Startdatum, B Starttag, C Startzeit, D Enddatum, E Endtag, F Endzeit, G Titel
EXPECTED_MIN_COLUMNS = 7


def excel_to_ics(excel_file_path, ics_file_path):
    """Liest die Excel-Terminliste ein und schreibt eine ICS-Datei.

    Gibt (anzahl_importierter_termine, liste_uebersprungener_zeilen) zurück.
    liste_uebersprungener_zeilen enthält (zeilennummer, grund)-Tupel.
    """
    wb = load_workbook(excel_file_path)
    sheet = wb.worksheets[0]

    # Punkt 9: Struktur vorab prüfen, statt erst beim Zugriff auf eine nicht
    # vorhandene Spalte mit einem IndexError abzustürzen.
    if sheet.max_column < EXPECTED_MIN_COLUMNS:
        raise ValueError(
            f'Die Excel-Datei hat nur {sheet.max_column} Spalte(n), '
            f'erwartet werden mindestens {EXPECTED_MIN_COLUMNS} '
            f'(A-G: Startdatum, Starttag, Startzeit, Enddatum, Endtag, Endzeit, Titel).'
        )

    cal = Calendar()
    cal.add('prodid', '-//HEMS SCHULJAHRESKALENDER//mxm.dk//')
    cal.add('version', '2.0')

    imported = 0
    skipped_rows = []

    for row_number, row in enumerate(sheet.iter_rows(min_row=4), start=4):
        # Komplett leere Zeilen einfach überspringen, ohne sie als Fehler zu zählen
        if all(cell.value is None for cell in row):
            continue

        start_raw = row[0].value
        if not isinstance(start_raw, datetime.datetime):
            skipped_rows.append((row_number, f'Ungültiges Startdatum: "{start_raw}"'))
            continue

        end_raw = row[3].value
        if end_raw is not None and not isinstance(end_raw, datetime.datetime):
            skipped_rows.append((row_number, f'Ungültiges Enddatum: "{end_raw}"'))
            continue

        summary_text = row[6].value
        if not summary_text:
            skipped_rows.append((row_number, 'Kein Titel (Spalte G) angegeben'))
            continue

        start_date_only = start_raw.date()
        end_date_only = end_raw.date() if end_raw is not None else None

        has_start_time = isinstance(row[2].value, datetime.time)
        has_end_time = isinstance(row[5].value, datetime.time)
        is_timed = has_start_time or has_end_time

        # Punkt 2+3: DTSTART/DTEND konsistent im selben Typ (date oder datetime)
        # erzeugen und immer ein DTEND setzen, statt es in Randfällen wegzulassen.
        if is_timed:
            dtstart = datetime.datetime.combine(
                start_date_only,
                row[2].value if has_start_time else datetime.time(0, 0)
            )
            if end_date_only is not None:
                if has_end_time:
                    dtend = datetime.datetime.combine(end_date_only, row[5].value)
                else:
                    # Enddatum ohne Uhrzeit -> bis Mitternacht des Folgetages
                    dtend = datetime.datetime.combine(
                        end_date_only + datetime.timedelta(days=1), datetime.time(0, 0)
                    )
            else:
                if has_end_time:
                    dtend = datetime.datetime.combine(start_date_only, row[5].value)
                else:
                    # Keine Endangabe vorhanden -> Termin ohne Dauer (Zeitpunkt)
                    dtend = dtstart
        else:
            dtstart = start_date_only
            if end_date_only is not None:
                dtend = end_date_only + datetime.timedelta(days=1)
            else:
                dtend = start_date_only + datetime.timedelta(days=1)

        event = Event()

        # Punkt 4: deterministische UID statt zufälliger UUID. So aktualisiert
        # ein erneuter Import bestehende Termine in Outlook, statt sie zu
        # duplizieren, solange sich Datum/Zeit/Titel nicht geändert haben.
        uid_source = f"{dtstart}|{dtend}|{summary_text}"
        uid_hash = hashlib.md5(uid_source.encode('utf-8')).hexdigest()
        event.add('uid', f"{uid_hash}@excel2ical")

        # Punkt 10: DTSTAMP gemäß RFC 5545 in UTC statt in lokaler Zeit ohne
        # Zeitzonenangabe.
        event.add('dtstamp', datetime.datetime.now(datetime.timezone.utc))

        event.add('summary', summary_text)
        event.add('dtstart', dtstart)
        event.add('dtend', dtend)

        # Spalte H (Index 7) mit einer separaten Beschreibung existiert in der
        # aktuellen Terminplan-Vorlage nicht mehr (nur A-G). Nur zugreifen,
        # wenn die Spalte tatsächlich vorhanden ist.
        if len(row) > 7 and row[7].value is not None:
            event.add('description', row[7].value)

        cal.add_component(event)
        imported += 1

    try:
        with open(ics_file_path, 'wb') as f:
            f.write(cal.to_ical())
    except Exception as e:
        raise IOError(f'ICS-Datei konnte nicht geschrieben werden: {e}')

    return imported, skipped_rows


def browse_excel_file():
    global exc_file_path
    path = filedialog.askopenfilename(
        title='Excel-Datei zur Konvertierung auswählen',
        filetypes=[('Excel Dokument', '*.xlsx')]
    )
    if not path:
        return
    exc_file_path = path
    filename = os.path.basename(exc_file_path)
    excel_file_label.config(text=filename)


def browse_ics_file():
    global ics_file_path
    path = filedialog.asksaveasfilename(
        title='ICS-Datei bestimmen',
        filetypes=[('ICS Dokument', '*.ics')],
        defaultextension='.ics'
    )
    if not path:
        return
    ics_file_path = path
    filename = os.path.basename(ics_file_path)
    ics_file_label.config(text=filename)


def convert_files():
    if not exc_file_path or not ics_file_path:
        messagebox.showerror(
            'Fehler',
            'Bitte sowohl eine Excel-Datei auswählen,\nals auch den Namen einer ICS-Datei bestimmen'
        )
        return

    try:
        imported, skipped_rows = excel_to_ics(exc_file_path, ics_file_path)
    except ValueError as e:
        messagebox.showerror('Fehler in der Excel-Struktur', str(e))
        return
    except IOError as e:
        messagebox.showerror('Fehler', str(e))
        return
    except Exception as e:
        messagebox.showerror('Unerwarteter Fehler', str(e))
        return

    # Punkt 6+7: eine gesammelte Rückmeldung am Ende statt eines Popups pro
    # fehlerhafter Zeile, inklusive Anzahl importierter Termine.
    message = f'Die Datei {ics_file_path}\nwurde erfolgreich erzeugt.\n\n{imported} Termin(e) importiert.'

    if skipped_rows:
        message += f'\n{len(skipped_rows)} Zeile(n) übersprungen:\n'
        max_shown = 15
        for row_number, reason in skipped_rows[:max_shown]:
            message += f'- Zeile {row_number}: {reason}\n'
        if len(skipped_rows) > max_shown:
            message += f'... und {len(skipped_rows) - max_shown} weitere.'
        messagebox.showwarning('Fertig (mit Hinweisen)', message)
    else:
        messagebox.showinfo('Erfolg', message)


# important for pyinstaller
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# create window
root = Tk()
root.iconbitmap(resource_path(WINDOW_ICON))
root.title(WINDOW_TITLE)

# Set the window not resizable
root.resizable(0, 0)

# create frames (for coloring background)
title_frame = Frame(root, bg=BG_COLOR)
title_frame.grid()
middle_frame = Frame(root)
middle_frame.grid()
bottom_frame = Frame(root, bg=BG_COLOR)
bottom_frame.grid()

# load image
pimage = PhotoImage(file=resource_path(MASCOT_PIC))
hems_logo = Label(title_frame, image=pimage, bg=BG_COLOR)
hems_logo.image = pimage

# create labels and buttons
description_label = Label(title_frame, text=APP_DESCRIPTION, font=FONT_DESCRIPTION, bg=BG_COLOR)
excel_file_label = Label(middle_frame, text=EXCEL_FILE_LABEL, font=FONT_LABEL)
browse_excel_button = Button(middle_frame, text=EXCEL_BUTTON_LABEL, command=browse_excel_file, font=FONT_BUTTONS)
ics_file_label = Label(middle_frame, text=ICS_FILE_LABEL, font=FONT_LABEL)
browse_ics_button = Button(middle_frame, text=ICS_BUTTON_LABEL, command=browse_ics_file, font=FONT_BUTTONS)
convert_button = Button(bottom_frame, text=CONVERT_BUTTON_TEXT, command=convert_files, font=FONT_CONVERT_BUTTON, bg=BG_COLOR_CONVERT_BUTTON)

# position labels, image and buttons
hems_logo.grid(row=0, column=0, padx=4, pady=10)
description_label.grid(row=0, column=1, pady=10)

excel_file_label.grid(row=2, column=0, columnspan=2)
browse_excel_button.grid(row=3, column=0, pady=10, columnspan=2)

ics_file_label.grid(row=4, column=0, columnspan=2)
browse_ics_button.grid(row=5, column=0, pady=10, columnspan=2)

convert_button.grid(row=7, column=0, padx=128, pady=10, columnspan=2)

root.mainloop()

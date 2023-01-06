import datetime
import os
import sys
import uuid
from openpyxl import load_workbook
from icalendar import Calendar, Event
from tkinter import BOTTOM, TOP, Button, Label, PhotoImage, Tk, filedialog, messagebox

def excel_to_ics(excel_file_path, ics_file_path):

    # Open the Excel file
    wb = load_workbook(excel_file_path)
    
    # Get the first sheet
    sheet = wb.worksheets[0]
    
    # Create a new ICS calendar
    cal = Calendar()
    
    # Set the calendar properties
    cal.add('prodid', '-//Der Schuljahreskalender//mxm.dk//')
    cal.add('version', '2.0')
    
    # Iterate over the rows in the sheet, starting from the second row, first row has the headlines
    for row in sheet.iter_rows(min_row=2):
        # Create a new ICS event
        event = Event()
 
        event.add('uid', uuid.uuid4()) # UID of the event, each event needs a distinguished UID
        event.add('dtstamp', datetime.datetime.now()) # Date of creation of the event
        # Set the event properties using the data from the Excel file
        event.add('summary', row[0].value)
         
        start = row[1].value # start Date of the event
        # Check if information in start date is really a date
        if type(start) is not datetime.datetime:
            messagebox.showerror('Error', f'{start} ist kein Startdatum! Excel überprüfen!')
        
        # strip the information in the cell of the time (unnecessary and looks ugly in ICS)
        start = start.date()
        
        # Combine date and time if time is given
        if type(row[2].value) == datetime.time:
            start = datetime.datetime.combine(row[1].value, row[2].value)
        event.add('dtstart', start)

                
        if(row[3].value is not None):
            end = row[3].value # end Date of the event
            
            # Check if information in end date is really a date
            if type(end) is not datetime.datetime:
                messagebox.showerror('Fehler', f'{end} ist kein Enddatum! Excel überprüfen!')
            
            # strip the information in the cell of the time (unnecessary)
            end = end.date()
            
            if type(row[4].value) == datetime.time: # check if end time is given
                end = datetime.datetime.combine(row[3].value, row[4].value)
            event.add('dtend', end)

        if(row[6].value is not None):
            event.add('description', row[6].value)
        if(row[7].value is not None):
            event.add('location', row[7].value)
        
        # Add the event to the calendar
        cal.add_component(event)
    # Write the calendar to the ICS file
        # Write the calendar to the ICS file
    with open(ics_file_path, 'wb') as f:
        f.write(cal.to_ical())

def browse_excel_file():
    global excfilepath 
    excfilepath = filedialog.askopenfilename(title='Excel-Datei zur Konvertierung auswählen', filetypes=[('Excel Dokument', '*.xlsx')]) 
    # next two lines to only show the filename on the label. The complete path would be too long to print.
    filename = os.path.basename(excfilepath)
    excel_file_label.config(text="\n" + filename)

def browse_ics_file():
    global icsfilepath 
    icsfilepath = filedialog.asksaveasfilename(title='ICS-Datei bestimmen', filetypes=[('ICS Dokument', '*.ics')], defaultextension='.ics') 
    # next two lines to only show the filename on the label. The complete path would be too long to print.
    filename = os.path.basename(icsfilepath)
    ics_file_label.config(text="\n" + filename)

def convert_files():
    if not excfilepath or not icsfilepath:
        messagebox.showerror('Fehler', 'Bitte sowohl eine Excel-Datei auswählen, als auch den Namen einer ICS-Datei bestimmen')
    else:
        excel_to_ics(excfilepath, icsfilepath)
        messagebox.showinfo('Erfolg', 'Umwandlung erfolgreich')

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

root = Tk()
root.iconbitmap(resource_path("excel2ics.ico"))
root.title('Excel2ical v2.0 (buc @ hems.de)')

# Set the window size
root.geometry("300x440")

#load image and set make it a bit transparent
pimage = PhotoImage(file=resource_path("cal.png"))
pimage.alpha = 128

#position image
label1 =Label(root, image=pimage)

# Add labels and buttons for the user to see the selected filenames
ics_file_label = Label(root, text="\nKeinen ICS-Dateinamen bestimmt", font=("Helvetica", 10))
browse_ics_button = Button(root, text="ICS-Dateinamen bestimmen", command=browse_ics_file, font=("Helvetica", 12))
excel_file_label = Label(root, text="\nKeine Excel-Datei ausgewählt", font=("Helvetica", 10))
browse_excel_button = Button(root, text="Excel-Datei auswählen", command=browse_excel_file, font=("Helvetica", 12))
convert_button = Button(root, text="Go!", command=convert_files, font=("Helvetica", 14),bg="#ee2724")
title_label = Label(root, text="Excel2ical v2.0", font=("Helvetica", 14))
titlesub_label = Label(root, text="Wandelt die Excelterminliste in eine\nOutlook-importierbare ICS-Datei um.\n", font=("Helvetica", 10))

label1.pack(side=TOP)
title_label.pack(side=TOP)
titlesub_label.pack(side=TOP)
excel_file_label.pack(side=TOP)
browse_excel_button.pack(side=TOP)
ics_file_label.pack(side=TOP)
browse_ics_button.pack(side=TOP)
convert_button.pack(side=BOTTOM,pady=10)

root.mainloop()


import os
import sys
import openpyxl
import datetime
import uuid
from icalendar import Calendar, Event
from PyQt5 import QtGui, QtCore, QtWidgets
from PyQt5.QtWidgets import QMessageBox


def excel_to_ics(excel_file_path, ics_file_path):
    # Open the Excel file
    wb = openpyxl.load_workbook(excel_file_path)
    # Get the first sheet
    sheet = wb.worksheets[0]
    # Create a new ICS calendar
    cal = Calendar()
    # Set the calendar properties
    cal.add('prodid', '-//Der Schuljahreskalender//mxm.dk//')
    cal.add('version', '2.0')
    # Iterate over the rows in the sheet, starting from the second row
    for row in sheet.iter_rows(min_row=2):
        # Create a new ICS event
        event = Event()
 
        event.add('uid', uuid.uuid4()) # UID of the event
        event.add('dtstamp', datetime.datetime.now()) # Date of creation
        # Set the event properties using the data from the Excel file
        event.add('summary', row[0].value)
         
        start = row[1].value # start Date of the event
        # Check if information in start date is really a date
        if type(start) is not datetime.datetime:
            QMessageBox.critical(None, 'Error', f'{start} ist kein Startdatum! Excel überprüfen!')
        # Combine date and time if time is given
        if type(row[2].value) == datetime.time:
            start = datetime.datetime.combine(row[1].value, row[2].value)
        event.add('dtstart', start)

                
        if(row[3].value is not None):
            end = row[3].value # end Date of the event
            # Check if information in end date is really a date
            if type(end) is not datetime.datetime:
                QMessageBox.critical(None, 'Error', f'{end} ist kein Enddatum! Excel überprüfen!')
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
    with open(ics_file_path, 'wb') as f:
        f.write(cal.to_ical())


class FileConverter(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        self.setGeometry(100, 100, 800, 600)
        self.setWindowTitle('Excel2ical von buc@hems.de')
        # Create a layout
        layout = QtWidgets.QVBoxLayout(self)
        # Create an explanation for the FileDialog
        label1 = QtWidgets.QLabel("Bitte die Termin-Exceldatei auswählen:")
        font = QtGui.QFont()
        font.setPointSize(14)
        label1.setFont(font)
        layout.addWidget(label1)
        # Create a file browser widget for selecting the Excel file
        self.excel_file_browser = QtWidgets.QFileDialog()
        self.excel_file_browser.setFileMode(QtWidgets.QFileDialog.ExistingFile)
        self.excel_file_browser.setNameFilter("Excel Datei mit den Terminen (*.xlsx)")
        # Create a line to seperate both file dialogs visually
        layout.addWidget(self.excel_file_browser)
        line1 = QtWidgets.QLabel()
        line1.setFixedHeight(2)
        line1.setStyleSheet("background-color: black;")
        layout.addWidget(line1)
        # Add another explanation line
        label2 = QtWidgets.QLabel("Bitte hier die ICS Datei auswählen, bzw. neu erstellen:")
        label2.setFont(font)
        layout.addWidget(label2)
        # Create another file browser widget for selecting the ICS file location
        self.ics_file_browser = QtWidgets.QFileDialog()
        self.ics_file_browser.setFileMode(QtWidgets.QFileDialog.AnyFile)
        self.ics_file_browser.setNameFilter("ICS Datei für den Import in Outlook (*.ics)")
        self.ics_file_browser.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
        layout.addWidget(self.ics_file_browser)
        #and another line
        line2 = QtWidgets.QLabel()
        line2.setFixedHeight(2)
        line2.setStyleSheet("background-color: black;")
        layout.addWidget(line2)
        # Create a button to start the conversion
        self.convert_button = QtWidgets.QPushButton("Konvertieren")
        # increase font size of button 
        fontbutton = QtGui.QFont()
        fontbutton.setPointSize(16)
        self.convert_button.setFont(fontbutton)
        self.convert_button.clicked.connect(self.convert)
        layout.addWidget(self.convert_button)
        
    def convert(self):
        # Get the selected Excel file path
        excel_file_path = self.excel_file_browser.selectedFiles()[0]
        # Get the selected ICS file path
        ics_file_path = self.ics_file_browser.selectedFiles()[0]
        # Ensure that the file has the .ics extension
        if not ics_file_path.endswith('.ics'):
                ics_file_path += '.ics'
        # Call the excel_to_ics function to convert the Excel data to an ICS file
        excel_to_ics(excel_file_path, ics_file_path)
        sys.exit()


if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    converter = FileConverter()
    converter.show()
    app.exec_()

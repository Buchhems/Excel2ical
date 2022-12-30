import os
import openpyxl
from icalendar import Calendar, Event
from PyQt5 import QtWidgets


def excel_to_ics(excel_file_path, ics_file_path):
    # Open the Excel file
    wb = openpyxl.load_workbook(excel_file_path)
    # Get the first sheet
    sheet = wb.worksheets[0]
    # Create a new ICS calendar
    cal = Calendar()
    # Iterate over the rows in the sheet, starting from the second row
    for row in sheet.iter_rows(min_row=2):
        # Create a new ICS event
        event = Event()
        # Set the event properties using the data from the Excel file
        event.add('summary', row[0].value)
        event.add('dtstart', row[1].value)
        event.add('dtend', row[4].value)
        event.add('description', row[5].value)
        event.add('location', row[6].value)
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
        # Create a layout
        layout = QtWidgets.QVBoxLayout(self)
        # Create a file browser widget for selecting the Excel file
        self.excel_file_browser = QtWidgets.QFileDialog()
        self.excel_file_browser.setFileMode(QtWidgets.QFileDialog.ExistingFile)
        self.excel_file_browser.setNameFilter("Excel files (*.xlsx)")
        layout.addWidget(self.excel_file_browser)
        # Create another file browser widget for selecting the ICS file location
        self.ics_file_browser = QtWidgets.QFileDialog()
        self.ics_file_browser.setFileMode(QtWidgets.QFileDialog.AnyFile)
        self.ics_file_browser.setNameFilter("ICS files (*.ics)")
        self.ics_file_browser.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
        layout.addWidget(self.ics_file_browser)
        # Create a button to start the conversion
        self.convert_button = QtWidgets.QPushButton("Convert")
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


if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    converter = FileConverter()
    converter.show()
    app.exec_()

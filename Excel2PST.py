import pandas as pd
import win32com.client

# Read the Excel file with the UTF-8 encoding, Excel sucks at importing umlauts. This is a safety margin.
df = pd.read_excel("path/to/file.xlsx", encoding="utf-8")

# Create a new PST file
outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
pst = mapi.AddStore("C:\\path\\to\\pst\\file.pst")

# Get the calendar folder of the PST file
calendar = pst.GetDefaultFolder(9)

# Iterate over the rows of the Excel file
for index, row in df.iterrows():
    # Get the values from the row
    topic = row['Betreff']
    start_date = row['Beginnt am']
    start_time = row['Beginnt um']
    end_date = row['Endet am']
    end_time = row['Endet um']
    description = row['Beschreibung']
    place = row['Ort']
    
    # Create a new calendar event
    event = outlook.CreateItem(1)
    event.Subject = topic
    event.Body = description
    event.Location = place
    
    # Check if start time and end time are specified, with this we can mark events without times as full day events.
    if pd.notnull(start_time) and pd.notnull(end_time):
        # Set the start and end time
        event.Start = start_date + " " + start_time
        event.End = end_date + " " + end_time
    else:
        # Set the event as an all-day event
        event.AllDayEvent = True
    
    # Save the event to the calendar
    event.Save()
    event.Move(calendar)

# Close the PST file
pst.Close()

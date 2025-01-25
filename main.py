import requests
import pandas as pd
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P


def save_to_sheet(file_path, df, sheet_name):
    # Load the existing ODS file
    doc = load(file_path)
    spreadsheet = doc.spreadsheet

    # Remove the sheet if it already exists
    for table in spreadsheet.getElementsByType(Table):
        if table.getAttribute("name") == sheet_name:
            spreadsheet.removeChild(table)

    # Create a new sheet
    new_table = Table(name=sheet_name)

    # Write column headers
    header_row = TableRow()
    for col in df.columns:
        cell = TableCell()
        cell.addElement(P(text=str(col)))
        header_row.addElement(cell)
    new_table.addElement(header_row)

    # Write data rows
    for _, row in df.iterrows():
        data_row = TableRow()
        for cell_value in row:
            cell = TableCell()
            cell.addElement(P(text=str(cell_value)))
            data_row.addElement(cell)
        new_table.addElement(data_row)

    # Append the new sheet to the spreadsheet
    spreadsheet.addElement(new_table)

    # Save the updated ODS file
    doc.save(file_path)


def upToDateCheck():
    response = requests.get('https://api.openf1.org/v1/meetings?meeting_key=latest')

    if response.status_code != 200:
        print(response.status_code)
        print("Server unresponsive!")
        quit()
    

    data = response.json()
    latest_key = data[0]['meeting_key']
    circ = data[0]['circuit_short_name']

    df = pd.read_excel('formule1 2025.ods', engine='odf', sheet_name='Races')
    sheet_key = df.loc[0, 'Last fetch']

    if latest_key == sheet_key:
        return 1
    else:
        return [latest_key, circ]


def getSessionPosition(driver_number, session_key):
    response = requests.get('https://api.openf1.org/v1/position?driver_number=' + str(driver_number) + '&session_key=' + str(session_key))
    positions = response.json()

    position = position[-1]['position']

    return position


def fetchData(upToDate):
    if upToDate == 1:
        print("Already up to date!")
        quit()

    meeting = upToDate[0]
    circ = upToDate[1]

    response = requests.get('https://api.openf1.org/v1/sessions?meeting_key=' + str(meeting) + '&session_name=Race')
    data = response.json()
    session_key = data[0]['session_key']
    
    df = pd.read_excel('formule1 2025.ods', engine='odf', sheet_name='Race positions')
    driver_list = df['Number']

    for driver in driver_list:
        df.loc[df['Number'] == driver, upToDate] = getSessionPosition(driver, session_key)
    
    print(df)



fetchData(upToDateCheck())
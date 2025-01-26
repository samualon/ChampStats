import requests
import pandas as pd
import time


def write_sheet(df, sheet, path):
    if df is not False:
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)


def upToDateCheck():
    response = requests.get('https://api.openf1.org/v1/meetings?meeting_key=latest')

    if response.status_code != 200:
        print(response.status_code)
        print("Server unresponsive! Try again later...")
        quit()
    

    data = response.json()
    latest_key = data[0]['meeting_key']
    circ = data[0]['circuit_short_name']

    df = pd.read_excel('formula1.xlsx', sheet_name='Races')
    sheet_key = df.loc[0, 'Last fetch']

    if latest_key == sheet_key:
        return 1
    else:
        df.loc[0, 'Last fetch'] = latest_key
        write_sheet(df, 'Races', 'formula1.xlsx')

        return [latest_key, circ]


def getSessionPosition(driver_number, session_key):
    response = requests.get('https://api.openf1.org/v1/position?driver_number=' + str(driver_number) + '&session_key=' + str(session_key))

    if response.status_code != 200:
        print(response.status_code)
        print("Server unresponsive! Try again later...")
        quit()

    print('https://api.openf1.org/v1/position?driver_number=' + str(driver_number) + '&session_key=' + str(session_key))
    positions = response.json()

    if len(positions) > 0:
        position = positions[-1]['position']
        return position


def fetchData(upToDate, session_name):
    if upToDate == 1:
        print("Already up to date!")
        quit()

    meeting = upToDate[0]
    circ = upToDate[1]

    response = requests.get('https://api.openf1.org/v1/sessions?meeting_key=' + str(meeting) + '&session_name=' + str(session_name))
    data = response.json()
    if len(data) > 0:
        session_key = data[0]['session_key']
    else:
        return False
    
    sheet_name = str(session_name) + ' positions'
    df = pd.read_excel('formula1.xlsx', sheet_name='Race positions')
    driver_list = df['Number']

    for driver in driver_list:
        df.loc[df['Number'] == driver, circ] = getSessionPosition(driver, session_key)
    
    return df


session_data = upToDateCheck()
write_sheet(fetchData(session_data, 'Sprint'), 'Sprint positions', 'formula1.xlsx')
time.sleep(5)
write_sheet(fetchData(session_data, 'Qualifying'), 'Qualifying positions', 'formula1.xlsx')
time.sleep(5)
write_sheet(fetchData(session_data, 'Race'), 'Race positions', 'formula1.xlsx')
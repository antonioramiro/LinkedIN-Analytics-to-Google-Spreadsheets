#Read README.md first

#--- Importing modules: ---#

import pandas as pd #manage data
import os #navigate directories
from datetime import date

#Interacting w/ sheets
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

print('1. DMT: all modules imported successfully.')

#-- Reading .xls files in directory: --#
def getSource():
    #Returns the names of the source spreadsheets, while checking if there are precisely 3 spreadsheets (visitors, followers and updates).
    parsedFilenames = []  # list of lists with the parsed .xls title, to verify if all 3 (visitor, follower and update data are present)
    source = [] #filenames of the source spreadsheets

    for file in os.listdir():
        if file.endswith(".xls"):
            parsedFilenames += [file.rsplit("_")]  
            source += [file]
    
    #verifies the number of spreadsheets in said directory
    if len(parsedFilenames) == 0:
        print('\nerror: No spreadsheet was found in this directory.\n')
    elif len(parsedFilenames) > 3:
        print('\nerror: There were more than 3 spreadsheets in this directory. There should only be one regarding visitors, one for followers data, and one for updates.\n')   
    elif len(parsedFilenames) < 3:
        print('\nerror: There were less than 3 spreadsheets in this directory. There should be one regarding visitors, one for followers data, and one for updates.\n')

    #verifies the existence of the three *diferent* spreadsheets
    elif sorted([filename[1] for filename in parsedFilenames]) != sorted(['visitors','followers','updates']): #verifies if the three spreadsheets are one of each kind
        print(f"\nerror: The files in the directory do not match the required. We found { sorted([filename[1] for filename in parsedFilenames])} instead of ['visitors', 'followers', 'updates']\n")
    else:
        print(f'3. Files found. Its names are {source[0]}, {source[1]} and {source[2]}.')
        return source

#Setting up gSheets: creating in the first run, or connecting to them, if it's not
def gSheetSetup():

    

    SCOPES = ['https://www.googleapis.com/auth/drive.file']  #permission  to only modify files created by the application

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    #find the web shortcuts in the folder, to determine whether this is the first run or not
    shortcuts = []
    for file in os.listdir():
        if file.endswith(".url"):
            shortcuts += [file]

    #if there is an incorrect number of .url shortcuts
    if len(shortcuts) > 1:
        print(f"\nerror: There should be one shortcut instead of {len(shortcuts)}.\n")

    #if there are no shortcuts, it means that it's the first run. Therefore:
    # + Spreadsheets have to be created and properly formated
    # + Shortcuts to said spreadsheets have to be created
    elif len(shortcuts) == 0:
        print("2. It's the first time you're running DMD. A new google spreadsheet wil be created on your GDrive's root. Changing its location won't affect this application")

        newpath = './archive' 
        if not os.path.exists(newpath):
            os.makedirs(newpath)

        # Creating spreadsheet       
        spreadsheet = { 
            "properties": { "title": 'Digital Marketing Dashboard', "timeZone": "Europe/London"},
            "sheets": [ 
                {"properties": { "title": 'Dashboard'} },
                {"properties": { "title": 'Value Center'} },
                {"properties": { "title": 'Updates'},
                    "data": [{ "startRow": 0, "startColumn": 0, "rowData": [ { "values": [ { "userEnteredValue": { "stringValue": 'UpdateTitle', } }, { "userEnteredValue": { "stringValue": 'UpdateLink', } }, { "userEnteredValue": { "stringValue": 'UpdateType', } }, { "userEnteredValue": { "stringValue": 'Campaign name', } }, { "userEnteredValue": { "stringValue": 'Posted by', } }, { "userEnteredValue": { "stringValue": 'Created date', } }, { "userEnteredValue": { "stringValue": 'Campaign start date', } }, { "userEnteredValue": { "stringValue": 'Campaign end date', } }, { "userEnteredValue": { "stringValue": 'Audience', } }, { "userEnteredValue": { "stringValue": 'Impressions', } }, { "userEnteredValue": { "stringValue": 'VideoViews', } }, { "userEnteredValue": { "stringValue": 'Clicks', } }, { "userEnteredValue": { "stringValue": 'CTR', } }, { "userEnteredValue": { "stringValue": 'Likes', } }, { "userEnteredValue": { "stringValue": 'Comments', } }, { "userEnteredValue": { "stringValue": 'Shares', } }, { "userEnteredValue": { "stringValue": 'Follows', } }, { "userEnteredValue": { "stringValue": 'EngagementRate', } }, { "userEnteredValue": { "stringValue": 'ApplauseRate', } } ] }, ], } ], },
                {"properties": { "title": 'Update Metrics', },
                    "data": [{ "startRow": 0, "startColumn": 0, "rowData": [ { "values": [ { "userEnteredValue": { "stringValue": 'Date', } }, { "userEnteredValue": { "stringValue": 'Impressions (organic)', } }, { "userEnteredValue": { "stringValue": 'Impressions (sponsored)', } }, { "userEnteredValue": { "stringValue": 'Impressions (total)', } }, { "userEnteredValue": { "stringValue": 'Unique impressions (organic)', } }, { "userEnteredValue": { "stringValue": 'Clicks (organic)', } }, { "userEnteredValue": { "stringValue": 'Clicks (sponsored)', } }, { "userEnteredValue": { "stringValue": 'Clicks (total)', } }, { "userEnteredValue": { "stringValue": 'Reactions (organic)', } }, { "userEnteredValue": { "stringValue": 'Reactions (sponsored)', } }, { "userEnteredValue": { "stringValue": 'Reactions (total)', } }, { "userEnteredValue": { "stringValue": 'Comments (organic)', } }, { "userEnteredValue": { "stringValue": 'Comments (sponsored)', } }, { "userEnteredValue": { "stringValue": 'Comments (total)', } }, { "userEnteredValue": { "stringValue": 'Shares (organic)', } }, { "userEnteredValue": { "stringValue": 'Shares (sponsored)', } }, { "userEnteredValue": { "stringValue": 'Shares (total)', } }, { "userEnteredValue": { "stringValue": 'Engagement rate (organic)', } }, { "userEnteredValue": { "stringValue": 'Engagement rate (sponsored)', } }, { "userEnteredValue": { "stringValue": 'Engagement rate (total)', } } ] }, ], } ], },
                {"properties": { "title": 'New followers', }, 
                    "data": [{ "startRow": 0, "startColumn": 0, "rowData": [ { "values": [ { "userEnteredValue": { "stringValue": 'Date', } }, { "userEnteredValue": { "stringValue": 'Sponsored followers', } }, { "userEnteredValue": { "stringValue": 'Organic followers', } }, { "userEnteredValue": { "stringValue": 'Total followers', } } ] }, ], } ], },
                {"properties": { "title": 'Visitor Metrics', }, 
                    "data": [{ "startRow": 0, "startColumn": 0, "rowData": [ { "values": [ { "userEnteredValue": { "stringValue": 'Date', } }, { "userEnteredValue": { "stringValue": 'Overview page views (desktop)', } }, { "userEnteredValue": { "stringValue": 'Overview page views (mobile)', } }, { "userEnteredValue": { "stringValue": 'Overview page views (total)', } }, { "userEnteredValue": { "stringValue": 'Overview unique visitors (desktop)', } }, { "userEnteredValue": { "stringValue": 'Overview unique visitors (mobile)', } }, { "userEnteredValue": { "stringValue": 'Overview unique visitors (total)', } }, { "userEnteredValue": { "stringValue": 'Life page views (desktop)', } }, { "userEnteredValue": { "stringValue": 'Life page views (mobile)', } }, { "userEnteredValue": { "stringValue": 'Life page views (total)', } }, { "userEnteredValue": { "stringValue": 'Life unique visitors (desktop)', } }, { "userEnteredValue": { "stringValue": 'Life unique visitors (mobile)', } }, { "userEnteredValue": { "stringValue": 'Life unique visitors (total)', } }, { "userEnteredValue": { "stringValue": 'Jobs page views (desktop)', } }, { "userEnteredValue": { "stringValue": 'Jobs page views (mobile)', } }, { "userEnteredValue": { "stringValue": 'Jobs page views (total)', } }, { "userEnteredValue": { "stringValue": 'Jobs unique visitors (desktop)', } }, { "userEnteredValue": { "stringValue": 'Jobs unique visitors (mobile)', } }, { "userEnteredValue": { "stringValue": 'Jobs unique visitors (total)', } }, { "userEnteredValue": { "stringValue": 'Total page views (desktop)', } }, { "userEnteredValue": { "stringValue": 'Total page views (mobile)', } }, { "userEnteredValue": { "stringValue": 'Total page views (total)', } }, { "userEnteredValue": { "stringValue": 'Total unique visitors (desktop)', } }, { "userEnteredValue": { "stringValue": 'Total unique visitors (mobile)', } }, { "userEnteredValue": { "stringValue": 'Total unique visitors (total)', } } ] }, ], } ], } ]}

        spreadsheet = service.spreadsheets().create(body=spreadsheet,fields='spreadsheetId').execute()
            
        # if needed, the following line prints the spreadsheet ID
        #  print('Spreadsheet ID: {0}'.format(spreadsheet.get('spreadsheetId')))    

        # Creating the shortcut to the spreadsheet           
        f = open("Digital Marketing Dashboard.url", "w")
        f.writelines(['[{000214A0-0000-0000-C000-000000000046}]','\nProp3=19,2','\n[InternetShortcut]','\nIDList=','\nURL=https://docs.google.com/spreadsheets/d/'+ spreadsheet.get('spreadsheetId') +'/'])
        f.close()
        

    
    shortcuts = []
    for file in os.listdir():
        if file.endswith(".url"):
            shortcuts += [file]

    #If the spreadsheets are already created, it's time to add the new data from the .xls files to them.
    if len(shortcuts) == 1:  
        dataSource = getSource()

        #Firstly, we have to get the ID of the spreadsheet, which is stored in the URL file and requires some parsing to obtain
        f = open("Digital Marketing Dashboard.url", "r")
        lines = f.readlines()
        spreadsheetID = lines[4].rsplit('/')[5]
        f.close()
        

        


        ##KPIs:
        #newValues = []

        ##applause rate
        #for row in values:
        #    print(row)
        #    newRow = []
        #    newRow = row + [int(row[13])/int(row[11])] + 
        #    newValues += newRow

        #values=newValues

        #with the IDs, the spreadsheets can now be opened and the data written
        body = {'values': pd.read_excel(dataSource[1],sheet_name='Update engagement',skiprows=[1,2],keep_default_na=False).values.tolist()}
        service.spreadsheets().values().append(spreadsheetId=spreadsheetID, range='Updates',valueInputOption='USER_ENTERED', body=body).execute()
        
        

        body = {'values': pd.read_excel(dataSource[0],sheet_name='New followers',skiprows=[1],keep_default_na=False).values.tolist()}
        service.spreadsheets().values().append(spreadsheetId=spreadsheetID, range='New followers',valueInputOption='USER_ENTERED', body=body).execute()        

        body = {'values': pd.read_excel(dataSource[2],sheet_name='Visitor metrics',skiprows=[1],keep_default_na=False).values.tolist()}
        service.spreadsheets().values().append(spreadsheetId=spreadsheetID, range='Visitor Metrics',valueInputOption='USER_ENTERED', body=body).execute()
        
        body = {'values': pd.read_excel(dataSource[1],sheet_name='Update metrics (aggregated)',skiprows=[1,2],keep_default_na=False).values.tolist()}
        service.spreadsheets().values().append(spreadsheetId=spreadsheetID, range='Update Metrics',valueInputOption='USER_ENTERED', body=body).execute()
       
        today = date.today().strftime("%Y-%m-%d")
        os.rename(dataSource[0], f"./archive/{today}_{dataSource[0]}")
        os.rename(dataSource[1], f"./archive/{today}_{dataSource[1]}")
        os.rename(dataSource[2], f"./archive/{today}_{dataSource[2]}")


gSheetSetup()   

        


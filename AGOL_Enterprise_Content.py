# Written by frederickkellner@usda.gov for questions please contact me.
# This script requires the ESRI Python API and Pandas
# these are not part of the standard Python 3 library but are availble in
# most installations of Arc Pro. They can also easily be added to stand alone
# Python 3 installation using Pip
# This script requires an excel spreadsheet with current FS employees.
# The best way to get this spreadsheet is to vist the following site
# http://fsweb.wo.fs.fed.us/hrm/data-reports/index.php#three
# then look for "ROSTER" link once you have the roster open you can filter
# by Region under the first column "Organization Breakdown"
# For example I primarily filter by "Region 01",
# I find it easier to copy these records and paste into a new excel spreadsheet
# including column headers and be sure not to change column headers
# MAKE SURE YOU SPREADSHEET has only 1 single sheet.
# Provide the path of the of your filtered roster to the "excelfile" variable


from arcgis.gis import GIS
import pandas as pd
import datetime
import os

# These variables that must be entered
username = 'MyAGOLUserName'  # your AGOL username
password = '123Iates0mehotdogz'  # your AGOL password
# provide the path to the excel file mentioned at the top of this script
excelFile = r'C:\path\to\your\excelfile'
# an output location for the output excel file of user content
outputFolder = r'C:\path\where\outputfile\will\be\written'


def computeDateFromESRITimestamp(timestampFromESRI):
    '''Takes epoch timestamp from ESRI and converst to Year-Month-Day'''
    d = datetime.datetime.fromtimestamp(
        timestampFromESRI/1000.0
        ).strftime('%Y-%m-%d')
    return d


def listStringJoiner(inputList, joiner=","):
    '''listStringJoiner(sequence[list]) -> string
    Takes an input list and returns a string where each
    element in the list is joined together to form a string
    returned string value. Default value to is to
    join with a comma.
    Example: [1,2,3,4] -> '1,2,3,4'
    '''
    stringJoinedList = joiner.join(
        str(itemFromList) for itemFromList in inputList
        )
    return stringJoinedList


def findAGOLUserContent(user):
    # users can have multiple folders with content within each folder.
    # If "else" is invoked users only have a main content
    # folder and no sub folders
    AGOLItems = []
    userIndex = emails[emails == user.email].index[0]
    if len(user.folders) > 1:
        for folder in user.folders:
            # this assumes no folders within folders. Only items in a folder
            if user.items(folder=folder, max_items=10000):
                folderItems = user.items(folder=folder, max_items=10000)
                for folderItem in folderItems:
                    AGOLItems.append(folderItem)
        allUserItems = AGOLItems
    else:
        allUserItems = user.items()

    enterpriseContent.append(
        [
            user.email,
            user.firstName,
            user.lastName,
            Forest[userIndex],
            District[userIndex],
            allUserItems
            ]
        )


def formatUserNameAndContent(userNameAndContent, usersContent):
    if usersContent.name is None:
        contentName = usersContent.title
    else:
        contentName = usersContent.name

    if usersContent.shared_with['groups']:
        sharedWith = listStringJoiner(usersContent.shared_with['groups'])
    else:
        sharedWith = 'Content is not shared with others'

    listForDataFrameConversion.append(
        [
            userNameAndContent[0],
            userNameAndContent[1],
            userNameAndContent[2],
            userNameAndContent[3],
            userNameAndContent[4],
            contentName,
            usersContent.type,
            sharedWith,
            computeDateFromESRITimestamp(usersContent.created),
            computeDateFromESRITimestamp(usersContent.modified)
        ]
    )


currentDate = datetime.date.today().strftime('%m%d%Y')

# sets up connection AGOL. Must connect with Admin account
gis = GIS("https://www.arcgis.com", username, password)

# pd.read_excel will throw error if there mutliple sheets
# you can specify with sheet_name='FS Roster' for example as an argument.
FSEmployees = pd.read_excel(excelFile)
emails = FSEmployees['WK_EMAIL_ADDRESS']
Forest = FSEmployees['ORG_CODE_LEVEL_3_DESCR']
District = FSEmployees['ORG_CODE_LEVEL_4_DESCR']

print("Looking through FS-AGOL User's Content")
enterpriseContent = []
for userEmail in emails:
    AGOLUserSearch = gis.users.search(userEmail, max_users=1)
    if AGOLUserSearch:
        user = AGOLUserSearch[0]
        findAGOLUserContent(user)

print("Finding Level of Sharing for each User's, this may take a while.")
listForDataFrameConversion = []
for userNameAndContent in enterpriseContent:
    for usersContent in userNameAndContent[5]:
        if usersContent.type != 'Service Definition':
            formatUserNameAndContent(userNameAndContent, usersContent)

excelFileName = 'AGOL_User_Content_{}.xlsx'.format(currentDate)

outputExcelPath = os.path.join(outputFolder, excelFileName)

finalSpreadhseet = pd.DataFrame(
    listForDataFrameConversion,
    columns=[
        'User Email',
        'First Name',
        'Last Name',
        'Forest',
        'Location',
        'Content',
        'Content Type',
        'Shared With',
        'Created Date',
        'Modified Date'
        ]
    )

finalSpreadhseet.to_excel(outputExcelPath, index=False)

print(
    'Script is complete check your results in the Excel file at',
    outputExcelPath
      )

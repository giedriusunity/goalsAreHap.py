import datetime
from urllib.request import urlopen
import urllib.parse
import xml.etree.ElementTree as ET
from datetime import *
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import json

#THIS IS HERE TO TEST IF THE COMMIT WORKS
#One more time
class Cases:
    def __init__(self, bugId, title, status, assignedTo, lastEdited):
        self.bugId = bugId
        self.title = title
        self.status = status
        self.assignedTo = assignedTo
        self.lastEdited = lastEdited


class Goals:
    def __init__(self, g1, g2, g3, g4, offender1, offender2, offender3, offender4):
        self.g1 = g1
        self.g2 = g2
        self.g3 = g3
        self.g4 = g4
        self.offender1 = offender1
        self.offender2 = offender2
        self.offender3 = offender3
        self.offender4 = offender4


def GoalNotMet(whichGoal, caseValues, data):
    for i in range(0, len(caseValues)):
        caseValues[i].lastEdited = str(caseValues[i].lastEdited)
        caseValues[i].lastEdited = caseValues[i].lastEdited[0:10]
    caseValues.sort(key=lambda x: x.lastEdited)

    print("Goal ", whichGoal, " not met")
    print("Starting updating botttleneck cases")  ##################################################################
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SAMPLE_SPREADSHEET_ID = data['spreadsheet']
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    # Call the Sheets API
    range_ = "'Goals'!A7:E999"

    print(
        "Uploading bottlenecks to data report sheet")  ##################################################################
    howManyCasesToShow = 10
    arrayOfValues = [Cases(None, None, None, None, None)] * howManyCasesToShow

    for i in range(0, howManyCasesToShow):
        if caseValues[i].bugId is not None:
            arrayOfValues[i] = Cases(caseValues[i].bugId,
                                     caseValues[i].title,
                                     caseValues[i].status,
                                     caseValues[i].assignedTo,
                                     caseValues[i].lastEdited)
        else:
            arrayOfValues[i] = Cases("", "", "", "", "")
    # print(arrayOfValues[0].bugId,arrayOfValues[0].title,arrayOfValues[0].status,arrayOfValues[0].assignedTo,arrayOfValues[0].lastEdited)
    values = {'values': []}
    for i in range(0, len(arrayOfValues)):
        values['values'].append([arrayOfValues[i].bugId,
                                 arrayOfValues[i].title,
                                 arrayOfValues[i].status,
                                 arrayOfValues[i].assignedTo,
                                 arrayOfValues[i].lastEdited])

    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()


def GoalCurrentValues(goal, data):
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SAMPLE_SPREADSHEET_ID = data['spreadsheet']
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    range_ = "'Goals'!F7:K10"

    # goal 2 is weird so its a little different
    # it will show how many cases above 24 there are
    goal2Correct = goal.offender2 - 24
    if goal2Correct < 0:
        goal2Correct = 0

    values = {
        'values': [
            ["Goal 1", goal.g1, "", "Goal 1", goal.offender1],
            ["Goal 2", goal.g2, "", "Goal 2", goal2Correct],
            ["Goal 3", goal.g3, "", "Goal 3", goal.offender3],
            ["Goal 4", goal.g4, "", "Goal 4", goal.offender4]
        ]
    }
    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()


##################################################################
def CheckGoals(data):
    caseBottleneckAmount = 0

    goalBottleneckReached = False
    offenderCount = 0
    goalValues = Goals(0, 0, 0, 0, 0, 0, 0, 0)

    # FIRST GOAL##################################################################
    token = data['token']
    print("Starting AssignedTo Query")  ###################################################################
    assignedFilter = data['assignedFilter']
    searchOldestUpdate = urllib.request.quote(assignedFilter)
    returnedOldestUpdate = "ixBug,sTitle,sStatus,sPersonAssignedTo,dtLastUpdated"
    queryOldestUpdate = "http://fogbugz.unity3d.com/api.asp?cmd=search&q=" + searchOldestUpdate + "&cols=" + returnedOldestUpdate + "&token=" + token
    response = urlopen(queryOldestUpdate)
    # print(queryOldestUpdate)
    print("Starting AssignedTo Decode")  #

    content = response.read().decode('utf-8')

    print("Starting Oldest Update Parsing")  ###################################################################
    tree = ET.ElementTree(ET.fromstring(content))
    oldestUpdateBugId = tree.findall(".//ixBug")
    oldestUpdateTitle = tree.findall(".//sTitle")
    oldestUpdateStatus = tree.findall(".//sStatus")
    oldestUpdateAssignedTo = tree.findall(".//sPersonAssignedTo")
    oldestUpdateLastEdited = tree.findall(".//dtLastUpdated")

    print("\nStarting Check if Goal 1 is Met")  ###################################################################
    offenderCases = [Cases(None, None, None, None, None)] * len(oldestUpdateLastEdited)
    caseValues = [0] * len(oldestUpdateBugId)
    for i in range(0, len(oldestUpdateBugId)):
        caseValues[i] = Cases(oldestUpdateBugId[i].text,
                              oldestUpdateTitle[i].text,
                              oldestUpdateStatus[i].text,
                              oldestUpdateAssignedTo[i].text,
                              oldestUpdateLastEdited[i].text)

    highest = 0
    goalIsMet = True
    goalCount = int(data['goalsInOrder'][0])
    goal = datetime.now() - timedelta(days=goalCount)
    for i in range(0, len(oldestUpdateLastEdited)):

        dateStripped = oldestUpdateLastEdited[i].text[0:10]
        dateFormatted = datetime.strptime(dateStripped, "%Y-%m-%d")
        if i == 0 or (dateFormatted - goal).days < highest:
            highest = (dateFormatted - goal).days

        if (dateFormatted - goal).days < 0:
            offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                 caseValues[i].title,
                                                 caseValues[i].status,
                                                 caseValues[i].assignedTo,
                                                 caseValues[i].lastEdited)
            offenderCount = offenderCount + 1
            goalIsMet = False

    goalValues.g1 = goalCount - highest
    if goalIsMet:
        print("Goal 1 is met")
    else:

        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(1, offenderCases, data)
            goalBottleneckReached = True

    # SECOND GOAL##################################################################
    print("\nStarting Check if Goal 2 is Met")  ###################################################################

    goalCount = int(data['goalsInOrder'][1])
    goalValues.offender1 = offenderCount
    offenderCount = 0
    activeCases = 0
    offenderCases = [Cases(None, None, None, None, None)] * len(oldestUpdateLastEdited)
    for i in range(0, len(oldestUpdateStatus)):
        if oldestUpdateStatus[i].text == "Active (New)":
            offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                 caseValues[i].title,
                                                 caseValues[i].status,
                                                 caseValues[i].assignedTo,
                                                 caseValues[i].lastEdited)
            offenderCount = offenderCount + 1
            activeCases = activeCases + 1

    print("Current Active new case count: " + str(activeCases))
    if activeCases <= goalCount:
        goalValues.g2 = activeCases
        print("Goal 2 is met")
    else:

        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(2, offenderCases, data)
            goalBottleneckReached = True

    # THIRD GOAL##################################################################
    goalValues.offender2 = offenderCount
    offenderCount = 0
    goalCount = int(data['goalsInOrder'][2])
    goal = datetime.now() - timedelta(days=goalCount)

    offenderCases = [Cases(None, None, None, None, None)] * len(oldestUpdateLastEdited)
    print("\nStarting Check if Goal 3 is Met")  ###################################################################
    for i in range(0, len(oldestUpdateLastEdited)):
        dateStripped = oldestUpdateLastEdited[i].text[0:10]
        dateFormatted = datetime.strptime(dateStripped, "%Y-%m-%d")
        if i == 0 or goalCount - (dateFormatted - goal).days > goalValues.g3:
            if "Resolved (Fixed)" not in oldestUpdateStatus[i].text and "Resolved (Completed)" not in \
                    oldestUpdateStatus[i].text:
                print((dateFormatted - goal).days)
                goalValues.g3 = goalCount - (dateFormatted - goal).days
        if (dateFormatted - goal).days < 0 and "Resolved (Fixed)" not in oldestUpdateStatus[
            i].text and "Resolved (Completed)" not in oldestUpdateStatus[i].text:
            goalIsMet = False
            goalValues.g3 = goalCount - (dateFormatted - goal).days
            offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                 caseValues[i].title,
                                                 caseValues[i].status,
                                                 caseValues[i].assignedTo,
                                                 caseValues[i].lastEdited)
            offenderCount = offenderCount + 1
    if goalIsMet:
        print("Goal 3 is met")
    else:

        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(3, offenderCases, data)
            goalBottleneckReached = True
    # return

    # FOURTH GOAL##################################################################
    print("Starting Oldest in filters Query")  ###################################################################
    goalValues.offender3 = offenderCount
    offenderCount = 0

    newFilters = data['newFilters']

    assignedTo = urllib.request.quote('assignedTo:"QA Incoming" AND status:active(new) AND(')

    fullFilterForCount = assignedTo
    for filterIndex in range(0, len(newFilters)):

        fullFilterForCount = fullFilterForCount + urllib.request.quote(
            newFilters[filterIndex]['filterSearch'].replace("'", '"'))
        if filterIndex < len(newFilters) - 1:
            fullFilterForCount = fullFilterForCount + "OR"

    returnedOldestFilters = "ixBug,sTitle,sStatus,sPersonAssignedTo,dtOpened"

    queryOldestFilters = "http://fogbugz.unity3d.com/api.asp?cmd=search&q=" + fullFilterForCount + ")" + "&cols=" + returnedOldestFilters + "&token=" + token
    print(queryOldestFilters)
    response = urlopen(queryOldestFilters)

    content = response.read().decode('utf-8')
    tree = ET.ElementTree(ET.fromstring(content))
    oldestAllFiltersBugId = tree.findall(".//ixBug")
    caseValues = [0] * (len(oldestAllFiltersBugId) + 10)
    print(len(caseValues))

    totalCases = -1

    for filterIndex in range(0, len(newFilters)):

        searchOldestFilters = assignedTo + urllib.request.quote(
            newFilters[filterIndex]['filterSearch'].replace("'", '"'))
        returnedOldestFilters = "ixBug,sTitle,sStatus,sPersonAssignedTo,dtOpened"
        queryOldestFilters = "http://fogbugz.unity3d.com/api.asp?cmd=search&q=" + searchOldestFilters + ")" + "&cols=" + returnedOldestFilters + "&token=" + token
        response = urlopen(queryOldestFilters)
        # print(queryOldestFilters)
        print("Starting Oldest In Filters Decode")  ###################################################################

        content = response.read().decode('utf-8')

        print("Starting Oldest Update Parsing")  ###################################################################
        tree = ET.ElementTree(ET.fromstring(content))
        oldestFiltersBugId = tree.findall(".//ixBug")
        oldestFiltersTitle = tree.findall(".//sTitle")
        oldestFiltersStatus = tree.findall(".//sStatus")
        #       oldestFiltersAssignedTo = tree.findall(".//sPersonAssignedTo")
        oldestFiltersLastEdited = tree.findall(".//dtOpened")
        ##################################################################
        print("\nStarting Check if Goal 4 is Met with filter #", filterIndex)  #

        print(len(oldestFiltersBugId))
        for i in range(0, len(oldestFiltersBugId)):
            totalCases = totalCases + 1
            caseValues[totalCases] = Cases(oldestFiltersBugId[i].text,
                                           oldestFiltersTitle[i].text,
                                           oldestFiltersStatus[i].text,
                                           newFilters[filterIndex]['filterName'],
                                           oldestFiltersLastEdited[i].text)

    offenderCases = [Cases(None, None, None, None, None)] * totalCases
    print("0")
    highest = 0
    goalCount = int(data['goalsInOrder'][3])
    goal = datetime.now() - timedelta(days=goalCount)

    print("1")
    for i in range(0, totalCases):
        dateStripped = caseValues[i].lastEdited[0:10]
        dateFormatted = datetime.strptime(dateStripped, "%Y-%m-%d")

        if (dateFormatted - goal).days < 0:
            offenderCount = offenderCount + 1
            goalIsMet = False
            offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                 caseValues[i].title,
                                                 caseValues[i].status,
                                                 caseValues[i].assignedTo,
                                                 caseValues[i].lastEdited)
            print(offenderCases[offenderCount].bugId)
        if i == 0 or (dateFormatted - goal).days <= highest:
            highest = (dateFormatted - goal).days
            goalValues.g4 = goalCount - (dateFormatted - goal).days
            print(goalValues.g4, offenderCases[offenderCount].bugId)

    if goalIsMet:
        print("Goal 4 is met")
    else:

        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(4, offenderCases, offenderCount, data)

    print("BottleNeck case ammount:", caseBottleneckAmount)
    if caseBottleneckAmount < 10:
        casesBeyondBottleneck = 10
        potentialOffenderCases = [Cases(None, None, None, None, None)] * casesBeyondBottleneck
        caseValues = caseValues[0:totalCases]
        caseValues.sort(key=lambda x: x.lastEdited)
        for index in range(0, len(potentialOffenderCases)):
            potentialOffenderCases[index] = caseValues[index]

    goalValues.offender4 = offenderCount

    if goalValues.offender3==0 and goalValues.offender2==0 and goalValues.offender1==0:
        filteredPotentialOffendrers = RecursiveRemoveDupes(offenderCases, potentialOffenderCases)
        FillWithGoal4(filteredPotentialOffendrers, caseBottleneckAmount, data, goalValues)
    else:
        FillWithGoal4(potentialOffenderCases, caseBottleneckAmount, data, goalValues)

    print("4")

    return goalValues


def RecursiveRemoveDupes(offenders, potentialOffenders):
    print("Removing duplicates")
    filteredCases = [Cases(0,0,0,0,0)] * len(potentialOffenders)
    for i in range(0, len(potentialOffenders)):
        for j in range(0, len(offenders)):
            if potentialOffenders[i].bugId == offenders[j].bugId:
                potentialOffenders.pop(i)
                filteredCases = RecursiveRemoveDupes(offenders, potentialOffenders)
                return filteredCases
    filteredCases = potentialOffenders
    return filteredCases
##################################################################


def FillWithGoal4(potentialOffenders, caseBottleneckAmount, data, goalValues):
    print(
        "Starting filling the rest of sheet with oldest filter cases")  ##################################################################
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SAMPLE_SPREADSHEET_ID = data['spreadsheet']
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    # Call the Sheets API

    print(
        "Uploading bottlenecks to data report sheet")  ##################################################################
    sheetOffset = 0
    if goalValues.offender1 > 0:
        sheetOffset = goalValues.offender1
    elif goalValues.offender3 > 0:
        sheetOffset = goalValues.offender3
    elif goalValues.offender4 > 0:
        sheetOffset = goalValues.offender4
    range_ = "'Goals'!A" + str(7 + sheetOffset) + ":E999"

    # print(arrayOfValues[0].bugId,arrayOfValues[0].title,arrayOfValues[0].status,arrayOfValues[0].assignedTo,arrayOfValues[0].lastEdited)
    values = {'values': []}
    for i in range(0, len(potentialOffenders)):
        print(potentialOffenders[i].bugId)
        values['values'].append([potentialOffenders[i].bugId,
                                 potentialOffenders[i].title,
                                 potentialOffenders[i].status,
                                 potentialOffenders[i].assignedTo,
                                 potentialOffenders[i].lastEdited[0:10]
                                 ]
                                )

    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()


##################################################################
def ReportGoals(goalValues, regularDay, downIndex, data):
    # regular day = 1 means it is friday, 2 is any other day
    print("Starting updating team goals")  ##################################################################
    print("Getting correct line")  ##################################################################
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SAMPLE_SPREADSHEET_ID = data['spreadsheet']
    SAMPLE_RANGE_NAME = "'Goals Data'!A2:A999"
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    newSTR = str(values[len(values) - 2])
    newSTR = newSTR.replace("['", "")
    newSTR = newSTR.replace("']", "")

    highestIndex = int(newSTR)
    row = str(highestIndex + 1 + regularDay - downIndex)
    rowvalue = highestIndex + regularDay - downIndex
    range_ = "'Goals Data'!A" + row + ":N" + row
    print("Uploading goals to data report sheet")  ##################################################################
    if regularDay == 2:
        rowvalue = "Current"
    values = {
        'values': [[str(rowvalue),
                    str(goalValues.g1), int(data['goalsInOrder'][0]),
                    str(goalValues.g1 <= int(data['goalsInOrder'][0])),
                    str(goalValues.g2), int(data['goalsInOrder'][1]),
                    str(goalValues.g2 <= int(data['goalsInOrder'][1])),
                    str(goalValues.g3), int(data['goalsInOrder'][2]),
                    str(goalValues.g3 <= int(data['goalsInOrder'][2])),
                    str(goalValues.g4), int(data['goalsInOrder'][3]),
                    str(goalValues.g4 <= int(data['goalsInOrder'][3])),
                    str(datetime.now())]],
    }
    # request = service.spreadsheets().values().update(SAMPLE_SPREADSHEET_ID, range, True, values)
    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()


def main():
    with open('keyValues.json', encoding="utf8") as f:
        data = json.load(f)
    loopCases = 10
    firstRun = False
    dateOther = datetime.now()
    oneSecondTracker = datetime.now()
    hasNoBeenLockedThisWeek = True
    downIndex = 1
    while True:
        curDate = datetime.now()

        if firstRun == False or ((curDate - dateOther).total_seconds() / 60) >= loopCases:
            goalValues = CheckGoals(data)

        if date.today().weekday() == int(data['updateTime'][0]) and hasNoBeenLockedThisWeek:
            if datetime.now().hour == int(data['updateTime'][1]):
                print("Today is the weekly cutoff")
                hasNoBeenLockedThisWeek = False
                ReportGoals(goalValues, 1, 0, data)
                downIndex = 0

        if firstRun == False or ((curDate - dateOther).total_seconds() / 60) >= loopCases:
            print("Looping update")
            dateOther = curDate
            ReportGoals(goalValues, 2, downIndex, data)
            GoalCurrentValues(goalValues, data)
            firstRun = True

        downIndex = 1  # motherfucker this was a bitch to figure our how to do

        if (datetime.now() - oneSecondTracker).total_seconds() >= 1:
            print("Seconds until next caseUpate:", loopCases * 60 - (curDate - dateOther).total_seconds())
            oneSecondTracker = datetime.now()
        if date.today().weekday() == 6:
            hasNoBeenLockedThisWeek = False


main()

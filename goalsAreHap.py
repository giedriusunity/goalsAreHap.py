import datetime
from urllib.request import urlopen
import urllib.parse
import xml.etree.ElementTree as ET
from datetime import *
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import json


class Cases:
    def __init__(self, bugId, title, status, assignedTo, lastEdited, milestone):
        self.bugId = bugId
        self.title = title
        self.status = status
        self.assignedTo = assignedTo
        self.lastEdited = lastEdited
        self.milestone = milestone


class Goals:
    def __init__(self, g1, g2, g3, g4, offender1, offender2, offender3, offender4, triage):
        self.g1 = g1
        self.g2 = g2
        self.g3 = g3
        self.g4 = g4
        self.offender1 = offender1
        self.offender2 = offender2
        self.offender3 = offender3
        self.offender4 = offender4
        self.triage = triage


class ZeroCases:
    def __init__(self, filterName, filter,
                 beyond):  # this is probably not the way assarys in python classes should be but hey, it works
        self.filterName = []
        self.filter = []
        self.beyond = beyond


class ZeroData:
    def __init__(self, cases, counterLast, counterAllTime):
        self.cases = []
        self.counterLast = counterLast
        self.counterAllTime = counterAllTime


def ReturnSorted(caseValues, goalValues):
    print("Sorting Cases")
    for i in range(0, len(caseValues)):
        caseValues[i].lastEdited = str(caseValues[i].lastEdited)
        caseValues[i].lastEdited = caseValues[i].lastEdited[0:10]
    print("triage", goalValues.triage)
    if goalValues.triage > 0:
        triageCases = [Cases(None, None, None, None, None, None)] * goalValues.triage
        for i in range(0, goalValues.triage - 1):
            triageCases[i] = caseValues[i]
            caseValues.pop(i)

    caseValues.sort(key=lambda x: x.lastEdited)
    if goalValues.triage > 0:
        for i in range(0, goalValues.triage - 1):
            caseValues.insert(0, triageCases[i])
    return caseValues


def GoalNotMet(whichGoal, caseValues, data):
    print("Goal ", whichGoal, " not met")

    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SAMPLE_SPREADSHEET_ID = data['spreadsheet']
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    # Call the Sheets API

    print("Clearing Sheet")
    clearValues = [Cases("", "", "", "", "", "")] * 900
    range_ = "'Goals'!A7:E999"
    values = {'values': []}
    for i in range(0, len(clearValues)):
        values['values'].append([clearValues[i].bugId,
                                 clearValues[i].title,
                                 clearValues[i].status,
                                 clearValues[i].assignedTo,
                                 clearValues[i].lastEdited])

    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()

    print("Starting updating botttleneck cases")  #####################################################
    print(
        "Uploading bottlenecks to data report sheet")  ################################################################
    howManyCasesToShow = 10
    arrayOfValues = [Cases(None, None, None, None, None, None)] * 900
    if len(caseValues) <= 10:
        howManyCasesToShow = len(caseValues)
    for i in range(0, howManyCasesToShow):
        if caseValues[i].bugId is not None:
            arrayOfValues[i] = Cases(caseValues[i].bugId,
                                     caseValues[i].title,
                                     caseValues[i].status,
                                     caseValues[i].assignedTo,
                                     caseValues[i].lastEdited,
                                     caseValues[i].milestone)
            if i == 0 and whichGoal == 4:
                arrayOfValues[i].assignedTo = FindTurn(arrayOfValues[i], data)
        else:
            arrayOfValues[i] = Cases("", "", "", "", "", "")
    # print(arrayOfValues[0].bugId,arrayOfValues[0].title,arrayOfValues[0].status,arrayOfValues[0].assignedTo,arrayOfValues[0].lastEdited)
    values = {'values': []}
    for i in range(0, len(arrayOfValues)):
        if arrayOfValues[i].bugId == None:
            arrayOfValues[i].bugId = ""
        values['values'].append(['=HYPERLINK("https://fogbugz.unity3d.com/f/cases/' + arrayOfValues[i].bugId +
                                 '";"' + arrayOfValues[i].bugId + '")',
                                 arrayOfValues[i].title,
                                 arrayOfValues[i].status,
                                 arrayOfValues[i].assignedTo,
                                 arrayOfValues[i].lastEdited])

    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()


def GoalCurrentValues(goal, data):
    print("Setting Goal Values in sheet")
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SAMPLE_SPREADSHEET_ID = data['spreadsheet']
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    range_ = "'Goals'!F7:K11"

    # goal 2 is weird so its a little different
    # it will show how many cases above 24 there are

    values = {
        'values': [
            ["Goal 1", goal.g1, "", "Goal 1", goal.offender1],
            ["Goal 2", goal.g2, "", "Goal 2", goal.offender2],
            ["Goal 3", goal.g3, "", "Goal 3", goal.offender3],
            ["Goal 4", goal.g4, "", "Goal 4", goal.offender4],
            ["Triage:", goal.triage]
        ]
    }
    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()


def CheckGoals(data):
    caseBottleneckAmount = 0

    goalBottleneckReached = False
    offenderCount = 0
    goalValues = Goals(0, 0, 0, 0, 0, 0, 0, 0, 0)

    # FIRST GOAL##################################################################
    token = data['token']
    print("Starting AssignedTo Query")
    assignedFilter = data['assignedFilter']  # Filter with all team members assigned to them
    searchOldestUpdate = urllib.request.quote(
        assignedFilter)  # Quotes, brackets and other signs need to be formatted for web queries otherwise they fail
    returnedOldestUpdate = "ixBug,sTitle,sStatus,sPersonAssignedTo,dtLastUpdated,ixFixFor"
    queryOldestUpdate = "http://fogbugz.unity3d.com/api.asp?cmd=search&q=" + searchOldestUpdate + "&cols=" + \
                        returnedOldestUpdate + "&token=" + token
    response = urlopen(queryOldestUpdate)
    print(queryOldestUpdate)
    print("Starting AssignedTo Decode")  #

    content = response.read().decode('utf-8')

    print("Starting Oldest Update Parsing")  ###################################################################
    tree = ET.ElementTree(ET.fromstring(
        content))  # There might be better plugins for decoding XMLs but this is inbuilt and works well enoguh
    oldestUpdateBugId = tree.findall(".//ixBug")
    oldestUpdateTitle = tree.findall(".//sTitle")
    oldestUpdateStatus = tree.findall(".//sStatus")
    oldestUpdateAssignedTo = tree.findall(".//sPersonAssignedTo")
    oldestUpdateLastEdited = tree.findall(".//dtLastUpdated")
    oldestUpdatedMilestone = tree.findall(
        ".//ixFixFor")  # this is the milestone for the case, i.e Triage, Undecided, 2018.3...

    print("\nStarting Check if Goal 1 is Met")  ###################################################################
    offenderCases = [Cases(None, None, None, None, None, None)] * len(
        oldestUpdateLastEdited)  # Array for cases that are bad
    caseValues = [0] * len(oldestUpdateBugId)  # Array for all cases returned
    for i in range(0, len(oldestUpdateBugId)):
        caseValues[i] = Cases(oldestUpdateBugId[i].text,
                              oldestUpdateTitle[i].text,
                              oldestUpdateStatus[i].text,
                              oldestUpdateAssignedTo[i].text,
                              oldestUpdateLastEdited[i].text,
                              oldestUpdatedMilestone[i].text)

    highest = 0
    goalIsMet = True  # Oh if only
    goalCount = int(data['goalsInOrder'][0])  # Reads the days count or whatever else to compare to
    goal = datetime.now() - timedelta(days=goalCount)  # Date generated from json file to compare to
    for i in range(0, len(oldestUpdateLastEdited)):

        dateStripped = caseValues[i].lastEdited[0:10]
        dateFormatted = datetime.strptime(dateStripped,
                                          "%Y-%m-%d")  # Date from fogbugz comes in a ISO 8601 Notation, needs to be fixed up for comparison
        if i == 0 or (dateFormatted - goal).days < highest:
            highest = (dateFormatted - goal).days

        if (dateFormatted - goal).days < 0:  # if case is older than that day it gets added to the array
            offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                 caseValues[i].title,
                                                 caseValues[i].status,
                                                 caseValues[i].assignedTo,
                                                 caseValues[i].lastEdited,
                                                 caseValues[i].milestone)
            offenderCount = offenderCount + 1  # Counting up cases for stats
            goalIsMet = False

    goalValues.g1 = goalCount - highest
    if goalIsMet:
        print("Goal 1 is met")
    else:

        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(1, ReturnSorted(offenderCases, goalValues), data)
            goalBottleneckReached = True

    goalValues.offender1 = offenderCount
    # SECOND GOAL##################################################################
    print("\nStarting Check if Goal 2 is Met")

    goalCount = int(data['goalsInOrder'][1])

    offenderCount = 0
    activeCases = 0
    offenderCases = [Cases(None, None, None, None, None, None)] * len(
        oldestUpdateLastEdited)  # Resetting bad case array
    for i in range(0, len(oldestUpdateStatus)):
        if oldestUpdateStatus[i].text == "Active (New)":
            offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                 caseValues[i].title,
                                                 caseValues[i].status,
                                                 caseValues[i].assignedTo,
                                                 caseValues[i].lastEdited,
                                                 caseValues[i].milestone)
            offenderCount = offenderCount + 1
            activeCases = activeCases + 1

    print("Current Active new case count: " + str(activeCases))
    goalValues.g2 = activeCases
    if activeCases <= goalCount:

        print("Goal 2 is met")
    else:  # Since goal 2 is different, if it is failed it sends the oldest assigned to sheet

        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(2, ReturnSorted(offenderCases, goalValues), data)
            goalBottleneckReached = True

    goalValues.offender2 = offenderCount - int(data["goalsInOrder"][1])
    if goalValues.offender2 < 0:
        goalValues.offender2 = 0

    # THIRD GOAL##################################################################

    offenderCount = 0
    goalCount = int(data['goalsInOrder'][2])
    goal = datetime.now() - timedelta(days=goalCount)
    goalCountBugs = 2
    goalBugs = datetime.now() - timedelta(days=goalCountBugs)
    print(goalBugs)

    offenderCases = [Cases(None, None, None, None, None, None)] * len(oldestUpdateLastEdited)
    print("\nStarting Check if Goal 3 is Met")  ###################################################################
    for i in range(0, len(oldestUpdateLastEdited)):
        dateStripped = caseValues[i].lastEdited[0:10]
        dateFormatted = datetime.strptime(dateStripped, "%Y-%m-%d")
        if i == 0 or goalCount - (
                dateFormatted - goal).days > goalValues.g3:  # This is for all cases that are old and not resovled
            if "Resolved (Fixed)" not in oldestUpdateStatus[i].text \
                    and "Resolved (Completed)" not in caseValues[i].status:
                goalValues.g3 = goalCount - (dateFormatted - goal).days
                #print("Goal 3 value", goalValues.g3, caseValues[i].bugId)
        if ((dateFormatted - goal).days < 0 and "Resolved (Fixed)" not in oldestUpdateStatus[i].text \
                    and "Resolved (Completed)" not in caseValues[i].status) \
                or ((dateFormatted - goalBugs).days < 0
                    and "Active" in caseValues[i].status
                    and "Active (New)" not in caseValues[i].status
                    and "Active (Pending Information)" not in caseValues[i].status
                    and "105" in caseValues[i].milestone):  # This is for cases that are returned as converted but triage back from devs. milestone = 105 is Triage
            goalIsMet = False

            if "(New)" not in caseValues[i].status and "(Pending Information)" not in caseValues[i].status and \
                    "Active" in caseValues[i].status and "105" in caseValues[i].milestone:
                # Place returned triage cases at the start of the array
                offenderCases.insert(0, Cases(
                    caseValues[i].bugId,
                    caseValues[i].title,
                    caseValues[i].status,
                    caseValues[i].assignedTo,
                    caseValues[i].lastEdited,
                    caseValues[i].milestone))
                goalValues.triage += 1
                print(caseValues[i].bugId)

            else:

                offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                     caseValues[i].title,
                                                     caseValues[i].status,
                                                     caseValues[i].assignedTo,
                                                     caseValues[i].lastEdited,
                                                     caseValues[i].milestone)
#                print(goalCount, (dateFormatted - goal).days, (caseValues[i].bugId,
#                                                     caseValues[i].title,
#                                                     caseValues[i].status,
#                                                     caseValues[i].assignedTo,
#                                                     caseValues[i].lastEdited,
#                                                     caseValues[i].milestone))



            offenderCount = offenderCount + 1
    if goalIsMet:
        print("Goal 3 is met", goalValues.g3)

    else:

        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(3, ReturnSorted(offenderCases, goalValues), data)
            goalBottleneckReached = True

    goalValues.offender3 = offenderCount

    # FOURTH GOAL##################################################################

    print("Starting Oldest in filters Query")

    offenderCount = 0

    newFilters = data['newFilters']

    assignedTo = \
        urllib.request.quote(
            'assignedTo:"QA Incoming" AND category:"qa incoming incident" and status:active AND(')  # First we need to search all cases through all filters in the team

    fullFilterForCount = assignedTo
    for filterIndex in range(0, len(newFilters)):
        # First we need to search all cases through all filters in the team so we would know the max number the array might be
        # Thus all filters are combined
        fullFilterForCount = fullFilterForCount + urllib.request.quote(
            newFilters[filterIndex]['filterSearch'].replace("'", '"'))
        if filterIndex < len(newFilters) - 1:
            fullFilterForCount = fullFilterForCount + "OR"

    returnedOldestFilters = "ixBug,sTitle,sStatus,sPersonAssignedTo,dtOpened,ixFixFor"

    queryOldestFilters = "http://fogbugz.unity3d.com/api.asp?cmd=search&q=" + fullFilterForCount + ")" + "&cols=" + \
                         returnedOldestFilters + "&token=" + token
    print(queryOldestFilters)
    response = urlopen(queryOldestFilters)

    content = response.read().decode('utf-8')
    tree = ET.ElementTree(ET.fromstring(content))
    oldestAllFiltersBugId = tree.findall(".//ixBug")  # Only one field is needed for this
    caseValues = [0] * (len(
        oldestAllFiltersBugId) + 999)  # Well this contradicts the comments above but honestly with how retarded fogbugz searches are this is better

    totalCases = 0

    zeroTracker = ZeroCases([], [], 0)

    for filterIndex in range(0, len(
            newFilters)):  # Again, there probably is a better way to make arrays inside custom classes
        zeroTracker.filterName.append(newFilters[filterIndex]["filterName"])
        zeroTracker.filter.append(0)

    zeroGoal = datetime.now() - timedelta(days=1)
    for filterIndex in range(0, len(newFilters)):

        print("\nStarting Query Check if Goal 4 is Met with filter #", filterIndex)
        searchOldestFilters = assignedTo + urllib.request.quote(
            newFilters[filterIndex]['filterSearch'].replace("'", '"'))
        returnedOldestFilters = "ixBug,sTitle,sStatus,sPersonAssignedTo,dtOpened,ixFixFor"  # Returned changes here to date opened
        queryOldestFilters = "http://fogbugz.unity3d.com/api.asp?cmd=search&q=" + searchOldestFilters + ")" + "&cols=" \
                             + returnedOldestFilters + "&token=" + token
        response = urlopen(queryOldestFilters)
        # print(filterIndex, " ", queryOldestFilter#s)
        print("Starting Oldest In Filters Decode")

        content = response.read().decode('utf-8')

        print("Starting Oldest Update Parsing")
        tree = ET.ElementTree(ET.fromstring(content))
        oldestFiltersBugId = tree.findall(".//ixBug")
        oldestFiltersTitle = tree.findall(".//sTitle")
        oldestFiltersStatus = tree.findall(".//sStatus")
        #       oldestFiltersAssignedTo = tree.findall(".//sPersonAssignedTo")
        oldestFiltersLastEdited = tree.findall(".//dtOpened")
        oldestFiltersMilestone = tree.findall(".//ixFixFor")
        ##################################################################

        for i in range(0, len(oldestFiltersBugId)):

            caseValues[totalCases] = Cases(oldestFiltersBugId[i].text,
                                           oldestFiltersTitle[i].text,
                                           oldestFiltersStatus[i].text,
                                           newFilters[filterIndex]['filterName'],
                                           oldestFiltersLastEdited[i].text,
                                           oldestFiltersMilestone[i].text)
            dateStripped = caseValues[totalCases].lastEdited[0:10]
            dateFormatted = datetime.strptime(dateStripped, "%Y-%m-%d")

            if (dateFormatted - zeroGoal).days < -1:
                zeroTracker.filter[filterIndex] = zeroTracker.filter[
                                                      filterIndex] + 1  # How many cases each filter older than 1 day has

            totalCases = totalCases + 1

    ZeroTracker(data, zeroTracker)

    offenderCases = [Cases(None, None, None, None, None, None)] * totalCases

    highest = 0
    goalCount = int(data['goalsInOrder'][3])
    goal = datetime.now() - timedelta(days=goalCount)

    for i in range(0, totalCases - 1):
        dateStripped = caseValues[i].lastEdited[0:10]
        dateFormatted = datetime.strptime(dateStripped, "%Y-%m-%d")

        if (dateFormatted - goal).days < 0:
            offenderCount = offenderCount + 1
            goalIsMet = False
            offenderCases[offenderCount] = Cases(caseValues[i].bugId,
                                                 caseValues[i].title,
                                                 caseValues[i].status,
                                                 caseValues[i].assignedTo,
                                                 caseValues[i].lastEdited,
                                                 caseValues[i].milestone)
        if i == 0 or goalCount - (dateFormatted - goal).days > highest:
            highest = goalCount - (dateFormatted - goal).days
            goalValues.g4 = highest

    if goalIsMet:
        print("Goal 4 is met")
    else:
        print("Goal 4 is NOT met")
        if not goalBottleneckReached:
            caseBottleneckAmount = offenderCount
            GoalNotMet(4, ReturnSorted(offenderCases, goalValues), data)

    print("BottleNeck case ammount:", caseBottleneckAmount)
    casesBeyondBottleneck = 10
    potentialOffenderCases = [Cases("", "", "", "", "",
                                    "")] * casesBeyondBottleneck  # If there are less that 10 cases as offenders, also populate sheet with oldest cases for Goal4

    if caseBottleneckAmount < 10 and totalCases != 0:
        if totalCases < 10:
            loopLen = totalCases  # If there are less than 10 cases in total in filters
        else:
            loopLen = len(potentialOffenderCases)
        caseValues = caseValues[0:totalCases]
        caseValues = ReturnSorted(caseValues, goalValues)
        for index in range(0, loopLen):
            potentialOffenderCases[index] = caseValues[index]
            if index == 0:
                potentialOffenderCases[index].assignedTo = FindTurn(potentialOffenderCases[index],
                                                                    # Find whose turn is it to take the next case bad for Goal 4
                                                                    data, )

    goalValues.offender4 = offenderCount
    if goalValues.offender3 == 0 and goalValues.offender2 == 0 and goalValues.offender1 == 0:
        filteredPotentialOffendrers = RecursiveRemoveDupes(offenderCases,
                                                           potentialOffenderCases)  # If a cases appears cross filter, e.g "XR" for iOS and VR

        print("Filling while all other goals are met")
        FillWithGoal4(filteredPotentialOffendrers, caseBottleneckAmount, data, goalValues)
    else:
        print("Filling while some previous goal is not met")
        FillWithGoal4(potentialOffenderCases, caseBottleneckAmount, data, goalValues)

    return goalValues


def FindTurn(original, data):
    # Original is the oldest Goal4 Case

    print("Finding the turn for the person who needs to take oldest filter case")
    with open('turn.json', encoding="utf8") as f:
        turn = json.load(f)
    # potentialOffenderCases[index].assignedTo + "(" + data['turnOrder'][0]['owners'][0] + ")"
    formatted = original.assignedTo

    turnData = {"turn": [], "lastCase": ""}  # Copy over data from json for modification

    for i in range(0, len(turn["turn"])):
        turnData["turn"].append(turn["turn"][i])
    turnData["lastCase"] = turn['lastCase']

    # reads filter Id and names that those filters are owned by. Note: NEEDS TO BE WRITTEN THE WAY IT IS IN FOGBUGZ
    for index in range(0, len(data["turnOrder"])):

        if original.assignedTo == data["turnOrder"][index]["filterName"]:

            turnOrder = turn["turn"][index]  # Whose turn is it anyway?
            print(index, turnOrder)
            formatted = original.assignedTo + "(" + data["turnOrder"][index]['owners'][
                int(turnOrder)] + ")"  # E.g iOS to iOS(John Johnson)

            if original.bugId != turn[
                "lastCase"]:  # Last one that was oldest is stored in a json and is used for comparison
                query = "case:" + turn["lastCase"]
                queryFull = "http://fogbugz.unity3d.com/api.asp?cmd=search&q=" + \
                            query + "&cols=sPersonAssignedTo&token=" + \
                            data["token"]
                response = urlopen(queryFull)  # We need to check if it was taken by the person whose turn it was
                content = response.read().decode('utf-8')
                tree = ET.ElementTree(ET.fromstring(content))
                assignedTo = tree.findall(".//sPersonAssignedTo")
                personTurn = assignedTo[0].text
                # Now this looks retarded due to constant parsing and converting but I swear it isn't
                # I just find it is better to store data as string and convert to int when needed
                if data["turnOrder"][index]["owners"][int(turnData["turn"][
                                                              index])] == personTurn:  # Is the person the last oldest case is assigned to the person whose turn it was. Check keyValue for structure
                    turnData["turn"][index] = str(int(turnData["turn"][index]) + 1)  # Increase turn
                if int(turnData["turn"][index]) == len(data["turnOrder"][index]['owners']):
                    turnData["turn"][index] = "0"  # reset turn order
                turnData["lastCase"] = original.bugId
                formatted = original.assignedTo + "(" + data["turnOrder"][index]['owners'][int(turnOrder)] + ")"
                with open('turn.json', 'w') as outfile:
                    json.dump(turnData, outfile)  # Save the turn change
    return formatted


def RecursiveRemoveDupes(offenders, potentialOffenders):
    print("Removing duplicates")
    # Since if it finds a dupe it pops, the array size must be lowered. At the end we have a nicely sized array with no dupes
    filteredCases = [Cases(0, 0, 0, 0, 0, 0)] * len(potentialOffenders)
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
        "Starting filling the rest of sheet with oldest filter cases")  ##############################################
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
        "Uploading bottlenecks to data report sheet")  ###############################################################
    sheetOffset = 0  # So it would fill after the real cases
    if goalValues.offender1 > 0:
        sheetOffset = goalValues.offender1
    elif goalValues.offender3 > 0:
        sheetOffset = goalValues.offender3
    elif goalValues.offender4 > 0:
        sheetOffset = goalValues.offender4
    range_ = "'Goals'!A" + str(7 + sheetOffset) + ":E999"

    values = {'values': []}
    for i in range(0, len(potentialOffenders)):

        if potentialOffenders[i].bugId is None:
            potentialOffenders[i].bugId = ""
        values['values'].append(['=HYPERLINK("https://fogbugz.unity3d.com/f/cases/' + potentialOffenders[i].bugId +
                                 '";"' + potentialOffenders[i].bugId + '")',
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
    # regular day = 1 means it is thursday, 2 is any other day
    print("Starting updating team goals")
    print("Getting correct line")
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

    newSTR = str(values[len(
        values) - 2])  # In the google sheet, current will be 1 less than row length, but highest number of "week" will be 2 less
    newSTR = newSTR.replace("['", "")
    newSTR = newSTR.replace("']", "")

    highestIndex = int(newSTR)
    row = str(highestIndex + 1 + regularDay - downIndex)

    # Thats why on one run we need to check and overwrite, and on another we need to find the next line to write to
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


def ZeroTracker(data, zeroTracker):
    print("Starting ZeroTracker")
    with open('zeroCases.json', encoding="utf8") as f:
        zeroCaseData = json.load(f)
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SAMPLE_SPREADSHEET_ID = data['spreadsheet']
    range_ = "'Goals'!$L$7:$M$14"
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    values = {'values': []}
    zeroCaseData = ZeroCumulativeTracker(zeroCaseData, zeroTracker, data)
    for i in range(0, len(zeroTracker.filter)):
        values['values'].append([zeroTracker.filterName[i],
                                 zeroTracker.filter[i]])
    values["values"].append([""])
    sumThisWeek = 0
    for i in range(0, len(zeroCaseData["casesSinceLastZero"])):
        sumThisWeek += int(zeroCaseData["casesSinceLastZero"][i])
    values["values"].append(["Cases done when no team cases older than a day"])
    values["values"].append(["Total", zeroCaseData["casesAllZero"]])
    values["values"].append(["Weekly Delta", sumThisWeek])

    # request = service.spreadsheets().values().update(SAMPLE_SPREADSHEET_ID, range, True, values)
    request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
                                                     valueInputOption='USER_ENTERED', body=values)
    request.execute()


def ZeroCumulativeTracker(zeroCaseData, zeroTracker, data):
    print("Starting check if each platform has 0 cases older than a day")
    if str(datetime.now())[0:10] != zeroCaseData["lastReset"] and datetime.now().weekday() == 4:
        for i in range(0, len(zeroCaseData["casesSinceLastZero"])):
            zeroCaseData["casesSinceLastZero"][i] = "0"
            zeroCaseData["lastReset"] = str(datetime.now())[0:10]
    queryStart = urllib.request.quote("editedBy:'QA Incoming' AND ")
    for i in range(0, len(zeroTracker.filter)):
        if zeroTracker.filter[i] == 0:  # Marks that that team has reached 0 cases older than a day
            if zeroCaseData["datesReached"][i] == "0":
                dateStripped = str(datetime.now())[0:10]
                dateStripped = datetime.strptime(dateStripped, '%Y-%m-%d').strftime(
                    '%m/%d/%Y')  # Fogbugz takes date in this format.
                zeroCaseData["datesReached"][i] = dateStripped

                timeFormat = "" + str(datetime.now().hour - 1) + ":" + str(
                    datetime.now().minute)  # its queries also run on UK time
                zeroCaseData["timesReached"][i] = timeFormat
            teamEdited = "("
            for y in range(0, len(zeroCaseData["filters"][i]["owners"])):
                if y < len(zeroCaseData["filters"][i]["owners"]) - 1:
                    teamEdited = teamEdited + "editedBy:'" + zeroCaseData["filters"][i]["owners"][
                        y] + "' OR "  # Combining names of people into one query
                else:
                    teamEdited = teamEdited + "editedBy:'" + zeroCaseData["filters"][i]["owners"][y] + "')"
            query = "https://fogbugz.unity3d.com/api.asp?cmd=search&q=" + queryStart + \
                    urllib.request.quote(teamEdited) + \
                    urllib.request.quote(" AND lastedited:'" + zeroCaseData["datesReached"][i] + " "
                                         + zeroCaseData["timesReached"][i] + "..'") + \
                    "&cols=ixBug&token=" + data["token"]

            response = urlopen(query)
            content = response.read().decode('utf-8')
            tree = ET.ElementTree(ET.fromstring(content))
            bugIDsXML = tree.findall(".//ixBug")
            bugIDs = []
            for k in range(0, len(bugIDsXML)):
                bugIDs.append(bugIDsXML[k].text)  # Creates array for found done cases

            zeroCaseData = CheckForAlreadyEdited(i, bugIDs, zeroCaseData)

        else:
            print("Team", zeroCaseData["filters"][i]["filterName"], "has cases older than a day, get on that!")
            zeroCaseData["datesReached"][i] = "0"
    with open('zeroCases.json', 'w', encoding='utf8') as outfile:
        json.dump(zeroCaseData, outfile)
    return zeroCaseData


def CheckForAlreadyEdited(index, bugIDs, cases):
    print("Checking if cases that were done since zero reached are not dupes")
    caseArray = cases["casesChecked"]
    for y in range(0, len(bugIDs)):
        notDupe = True
        for i in range(0, len(cases[
                                  "casesChecked"])):  # Checks through all cases done up to this point. Now this can be a bit wasteful as it checks each case against 25k+ cases so the array might need to be reset after some time

            if bugIDs[y] == cases["casesChecked"][i]:
                notDupe = False
        if notDupe:
            caseArray.append(bugIDs[y])
            cases["casesSinceLastZero"][index] = str(int(cases["casesSinceLastZero"][index]) + 1)
            cases["casesAllZero"] = str(int(cases["casesAllZero"]) + 1)
    return cases


def main():
    with open('keyValues.json', encoding="utf8") as f:
        data = json.load(f)
    loopCases = 5
    firstRun = False
    dateOther = datetime.now()
    oneSecondTracker = datetime.now()
    hasNoBeenLockedThisWeek = True  # This would work in theory to auto lock each week but doesn't
    downIndex = 1
    print("Starting loop")
    while True:
        curDate = datetime.now()
        if 7 <= curDate.hour < 23:  # Runs only during work hours to not put pressure on fogbugz when not necessary

            if firstRun == False or ((curDate - dateOther).total_seconds() / 60) >= loopCases:
                goalValues = CheckGoals(data)

            if date.today().weekday() == int(data['updateTime'][
                                                 0]) and hasNoBeenLockedThisWeek:  # This activates when a week has passed and needs to move to another line
                if datetime.now().hour == int(data['updateTime'][1]):
                    print("Today is the weekly cutoff")
                    hasNoBeenLockedThisWeek = False
                    ReportGoals(goalValues, 1, 0, data)
                    downIndex = 0

            if firstRun is False or ((curDate - dateOther).total_seconds() / 60) >= loopCases:
                print("Looping update")
                dateOther = curDate
                ReportGoals(goalValues, 2, downIndex, data)
                GoalCurrentValues(goalValues, data)
                firstRun = True

            downIndex = 1  # motherfucker this was so simple a bitch to figure our how to do

            if (datetime.now() - oneSecondTracker).total_seconds() >= 1:
                print(data["team"] + ": Seconds until next caseUpate:",
                      loopCases * 60 - (curDate - dateOther).total_seconds())
                oneSecondTracker = datetime.now()
            if date.today().weekday() == 6:
                hasNoBeenLockedThisWeek = False
        else:
            print("Time of work is over, what are you still doing here?")


main()
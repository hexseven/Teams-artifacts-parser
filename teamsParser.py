#!/usr/bin/env python3
import json
from datetime import datetime, timezone
import argparse
import sys
import os


# file handle method
def fileHandler(filename):
    with open(filename, 'r') as file:
        data = json.load(file)  # Load the JSON data
        return data


# support methods
# Convert from Epoch time to human-readable
def convertFromEpochTime(time):
    humanReadableTime = float(time)
    humanReadableTime = (datetime.fromtimestamp(humanReadableTime / 1000.0)
                         .strftime('%Y-%m-%d %H:%M:%S.%f'))
    return humanReadableTime


# Convert from ISO time to human-readable
def convertFromIsoTime(time):
    humanReadableTime = time
    humanReadableTime = humanReadableTime.split('.')[0]
    humanReadableTime = datetime.strptime(humanReadableTime, '%Y-%m-%dT%H:%M:%S')
    return humanReadableTime


# Calculation call length method
def callLength(callStart, callEnd):
    callStart = callStart.split('.')[0]
    callEnd = callEnd.split('.')[0]
    callStart = datetime.strptime(callStart, '%Y-%m-%dT%H:%M:%S')
    callEnd = datetime.strptime(callEnd, '%Y-%m-%dT%H:%M:%S')
    timeDifference = callEnd - callStart
    return timeDifference


# Convert user email/mri to displayName method
def getUsername(email, data):
    email = email.split('.')[-1]
    displayName = ''
    for row in range(len(data)):
        if data[row]['record_type'] == 'contact':
            mri = data[row]['mri'].split('.')[-1]
            if mri == email:
                displayName = data[row]['displayName']
    return displayName


# Retrieve methods
# Return users from teams method
def getUsers(data):
    uniqueUsers = []  # unique contacts
    seenMris = set()  # duplicated mris

    for row in range(len(data)):
        currentUser = data[row]
        if currentUser['record_type'] == 'contact':
            temp = {}
            mri = currentUser['mri']
            displayName = currentUser['displayName']
            if mri not in seenMris:
                temp['displayName'] = currentUser['displayName']
                temp['mri'] = currentUser['mri']
                temp['email'] = currentUser['email']
                temp['userPrincipalName'] = currentUser['userPrincipalName']
                uniqueUsers.append(temp)
                seenMris.add(mri)
    return uniqueUsers


# Return message content method
def getMessageContent(data):
    listOfMessages = []
    for row in range(len(data)):
        temp = {}
        if data[row]['record_type'] == 'message':
            messageProperties = data[row]['properties']
            if not 'meeting' in messageProperties:
                temp['conversationId'] = data[row]['conversationId']
                temp['content'] = data[row]['content']
                temp['creator'] = getUsername(data[row]['creator'], data)
                temp['sent time'] = convertFromEpochTime(data[row]['originalArrivalTime'])
                fileBaseUrl = ''
                fileName = ''
                fileType = ''
                fileChicletState = ''
                filePreview = ''
                if 'files' in messageProperties.keys():
                    for property in messageProperties.values():
                        if property:
                            for item in property:
                                fileBaseUrl = item.get('baseUrl')
                                fileName = item.get('fileName')
                                fileType = item.get('fileType')
                                fileChicletState = item.get('fileChicletState')
                                filePreview = item.get('filePreview')
                if fileName:
                    temp['fileBaseUrl'] = fileBaseUrl
                    temp['fileName'] = fileName
                    temp['fileType'] = fileType
                    temp['fileChicletState'] = fileChicletState
                    temp['filePreview'] = filePreview
                    temp['cacheViewerId'] = filePreview.get('previewUrl').split("/")[-3]
                if 'emotions' in messageProperties.keys():
                    temp['emotions'] = messageProperties.get('emotions')
                listOfMessages.append(temp)
    return listOfMessages


# Return teams reactions method
def getReactions(data):
    listOfReactions = []
    for row in range(len(data)):
        temp = {}
        if data[row]['record_type'] == 'reaction':
            activityProperties = data[row].get('properties').get('activity')
            activityTimestamp = activityProperties.get('activityTimestamp')
            activityTimestamp = activityTimestamp.split('.')[0]
            activityTimestamp = datetime.strptime(activityTimestamp, '%Y-%m-%dT%H:%M:%S')
            clientArrivalTime = float(data[row].get('clientArrivalTime'))
            prettyClientArrivalTime = (datetime.fromtimestamp(clientArrivalTime / 1000.0)
                                       .strftime('%Y-%m-%d %H:%M:%S.%f'))
            temp['clientArrivalTime'] = prettyClientArrivalTime
            originalArrivalTime = float(data[row].get('originalArrivalTime'))
            prettyOriginalArrivalTime = (datetime.fromtimestamp(originalArrivalTime / 1000.0)
                                         .strftime('%Y-%m-%d %H:%M:%S.%f'))
            temp['originalArrivalTime'] = prettyOriginalArrivalTime
            temp['conversationId'] = data[row]['conversationId']
            temp['creator'] = getUsername(data[row]['creator'], data)
            temp['isFromMe'] = data[row]['isFromMe']
            temp['messageType'] = data[row]['messagetype']
            temp['activitySubtype'] = activityProperties.get('activitySubtype')
            temp['activityTimestamp'] = activityTimestamp
            temp['count'] = activityProperties.get('count')
            temp['messagePreview'] = activityProperties.get('messagePreview')
            temp['sourceThreadTopic'] = activityProperties.get('sourceThreadTopic')
            temp['sourceUserId'] = getUsername(activityProperties.get('sourceUserId'), data)
            temp['targetUserId'] = getUsername(activityProperties.get('targetUserId'), data)
            listOfReactions.append(temp)
    return listOfReactions


# Return call information's method
def getCalls(data):
    listOfCalls = []
    for row in range(len(data)):
        temp = {}
        if data[row]['record_type'] == 'call':
            callProperties = data[row]['properties']['call-log']
            temp['callDirection'] = callProperties['callDirection']
            temp['creator'] = getUsername(data[row]['creator'], data)
            temp['callType'] = callProperties['callType']
            temp['callState'] = callProperties['callState']
            temp['target'] = getUsername(callProperties['target'], data)
            callStart = callProperties['connectTime']
            callEnd = callProperties['endTime']
            temp['connectTime'] = convertFromIsoTime(callStart)
            temp['endTime'] = convertFromIsoTime(callEnd)
            temp['callLength'] = callLength(callStart, callEnd)
            listOfCalls.append(temp)
    return listOfCalls


# return meetings method
def getMeetings(data):
    listOfMeetings = []
    for row in range(len(data)):
        temp = {}
        if data[row].get('type') == 'Meeting':
            meetingProperties = data[row]['threadProperties']['meeting']
            threadProperties = data[row]['threadProperties']
            temp['meetingSubject'] = meetingProperties.get('subject')
            membersDict = data[row]['members']
            membersTable = []
            for user in membersDict:
                membersTable.append(getUsername(user.get('id'), data))
            temp['members'] = membersTable
            temp['creator'] = getUsername(threadProperties.get('creator'), data)
            temp['hasTranscript'] = threadProperties.get('hasTranscript')
            temp['isLastMessageFromMe'] = threadProperties.get('isLastMessageFromMe')
            temp['isLiveChatEnabled'] = threadProperties.get('isLiveChatEnabled')
            temp['meetingEndTime'] = meetingProperties.get('endTime')
            temp['meetingStartTime'] = meetingProperties.get('startTime')
            temp['meetingIsCancelled'] = meetingProperties.get('isCancelled')
            temp['meetingLocation'] = meetingProperties.get('location')
            temp['meetingURL'] = meetingProperties.get('meetingJoinUrl')
            temp['meetingType'] = meetingProperties.get('meetingType')
            temp['meetingOrganizerId'] = threadProperties.get('organizedId')
            temp['meetingScenario'] = meetingProperties.get('Scenario')
            temp['picture'] = threadProperties.get('picture')
            listOfMeetings.append(temp)
    return listOfMeetings


# Printing methods
# Create pretty console output
def printConsolePrettyOutput(dictionaryList):
    i = 1
    print(f'Count of items = {len(dictionaryList)}')
    for row in range(len(dictionaryList)):
        print('+' + '-' * 40)
        print(f'Item = {i}/{len(dictionaryList)}')
        for key, value in dictionaryList[row].items():
            print(f'[*] {key}: {value}')
        i += 1


# Print messages in particular threads
def printConsoleMessageThreads(dictionaryList):
    listOfThreads = []
    dictOfMessages = {}

    # retrieve list of thread id
    for thread in range(len(dictionaryList)):
        conversationId = dictionaryList[thread].get('conversationId')
        if conversationId not in listOfThreads:
            listOfThreads.append(conversationId)

    # create dictionary with messages grouped by conversation id {conversationId: [message, message]}
    for thread in listOfThreads:
        tableOfMessages = []
        counter = 1
        for row in range(len(dictionaryList)):
            conversationId = dictionaryList[row].get('conversationId')
            if thread == conversationId:
                tableOfMessages.append(dictionaryList[row])
            counter += 1
        dictOfMessages[thread] = tableOfMessages

    print(f'[INFO] Count of messages = {len(dictionaryList)}')
    print(f'[INFO] Count of threads = {len(listOfThreads)}')

    for thread in dictOfMessages:
        print('+' + '-' * (len(thread) + 2) + '+')
        print(f'| {thread} |')
        print('+' + '-' * (len(thread) + 2) + '+')
        print(f'[INFO] Count of messages in thread = {len(dictOfMessages.get(thread))}')
        usersInThread = []
        for message in dictOfMessages.get(thread):
            for key, value in message.items():
                if key == 'creator':
                    usersInThread.append(value)
        usersInThread = set(usersInThread)
        usersStr = ''
        for user in usersInThread: usersStr += user + ' '
        print(f'[INFO] Users in thread: {usersStr}\n')
        for message in dictOfMessages.get(thread):
            for key, value in message.items():
                if key != 'conversationId':
                    # print emotions
                    if key == 'emotions':
                        print(f'[*] emotions:')
                        for item in value:  # fixmes
                            print('-' * 50)
                            if not item.get('users'):
                                print(f'[!] Empty row. Previously was {item.get('key')},'
                                      f' but somebody changed emotion.')
                                print('')
                            for row in item.get('users'):
                                print(f'\t\temotion: {item.get('key')}')
                                for emotionKey, emotionValue in row.items():
                                    if emotionKey == 'time':
                                        print(f'\t\t{emotionKey}: {convertFromEpochTime(emotionValue)}')
                                    else:
                                        print(f'\t\t{emotionKey}: {emotionValue}')
                    # print rest
                    else:
                        print(f'[*] {key}: {value}')
            print('\n')


def printHeader(text):
    x = text.center(48, " ")
    print(f'\n\n+{'=' * 48}+')
    print(f'|{x}|')
    print(f'+{'=' * 48}+')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--file', help='Input file for analyze.')
    parser.add_argument('-o', '--output', default='output.txt', help='Name of output file. '
                                                                     'By default is "output.txt"')
    parser.add_argument('-u', '--users', action='store_true', help='Get users present in Teams.')
    parser.add_argument('-c', '--calls', action='store_true', help='Get information about calls.')
    parser.add_argument('-m', '--messages', action='store_true', help='Get message content.')
    parser.add_argument('-t', '--meetings', action='store_true', help='et meeting information.')
    parser.add_argument('-r', '--reactions', action='store_true', help='Get reactions data.')
    args = parser.parse_args()

    # handle input file
    inputFile = args.file
    data = fileHandler(inputFile)

    # handle output file
    outputFile = args.output
    outputFile = outputFile.replace('"', '').replace("'", "")
    if args.output:
        isExtension = os.path.splitext(args.output)[1]
        if isExtension.lower() == '.txt':
            pass
        else:
            outputFile += '.txt'

    # start creating output
    sys.stdout = open(outputFile, 'a+')
    print(f'Script execution time: {datetime.now(timezone.utc)} UTC')
    print(f'Script running for file {inputFile}')
    print('Script author: Patryk \'Hex7\' Łabuz')
    # handle arguments
    if args.users:
        printHeader('users')
        users = getUsers(data)
        printConsolePrettyOutput(users)
    if args.calls:
        printHeader('calls')
        calls = getCalls(data)
        printConsolePrettyOutput(calls)
    if args.messages:
        printHeader('messages')
        messages = getMessageContent(data)
        printConsoleMessageThreads(messages)
    if args.meetings:
        printHeader('meetings')
        meetings = getMeetings(data)
        printConsolePrettyOutput(meetings)
    if args.reactions:
        printHeader('reactions')
        reactions = getReactions(data)
        printConsolePrettyOutput(reactions)

    sys.stdout.close()


if __name__ == '__main__':
    main()

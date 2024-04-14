## Input File Handle
`def fileHandler(filename)`
</br>Simple method for file handling in read mode.
</br></br>**Parameter**</br>
filename : str 
</br>&emsp;&emsp;Path to input json file. Expected .json as output from ms_teams_parser.exe.
</br></br>**Returns**</br>
data : JSON object
</br>&emsp;&emsp;JSON ready for analysis

---
## Support Methods 

`def convertFromEpochTime(time)`
</br>Converts Epoch time to a human-readable format. The function takes an Epoch time value as input (provided as a string) and returns a formatted string representing the corresponding date and time.
</br></br>**Parameter**
</br>time : str 
</br>&emsp;&emsp;Receives Epoch time as string. 
</br></br>**Returns**</br>
humanReadableTime : str
</br>&emsp;&emsp;Human-readable time format. 

</br>

`callLength(callStart, callEnd)`
</br>Calculate time length based on start and end of call timestamps.
</br></br>**Parameters**</br>
callStart : str 
</br>&emsp;&emsp;The start of call timestamp in ISO 8601.
callEnd : str
</br>&emsp;&emsp;The end of call timestamp in ISO 8601.
</br></br>**Returns**</br>
timeDifference : str
</br>&emsp;&emsp;Timedelta as difference between callEnd and callStart.

</br>

`def getUsername(email, data)`
</br>Matching MRI with username of Teams user. Function looking for field record_type which value is contact. If matches, going deeper for displayName. 
Example:
```
"displayName": "Paulie Gualtieri",
"email": "live:.cid.2627cc1c36866fbf",
"mri": "8:live:.cid.2627cc1c36866fbf",
```
</br></br>**Parameters**</br>
email : str 
</br>&emsp;&emsp;MRI(Microsoft Resource Identifier) - unique id for every user in resource. 
data : JSON object
</br>&emsp;&emsp;JSON ready for analysis, output of fileHandler().
</br></br>**Returns**</br>
displayName : str
</br>&emsp;&emsp;Value of displayName key. Standar username e.g. Silvio Dante. 

---
## Analysis Methods

`def getUsers(data)`
</br>Retrieve contacts presents in Teams.
</br></br>**Parameter**</br>
data : JSON object
</br>&emsp;&emsp;JSON ready for analysis
</br></br>**Returns**</br>
uniqueUsers : list
</br>&emsp;&emsp;List of unique users. Each field in list contains dictionary with user attributes. Structure: `[{}, {}, ..., {}]`.

</br>

`def getMessageContent(data)`
</br>Retrieve messages from chats.
</br></br>**Parameter**</br>
data : JSON object
</br>&emsp;&emsp;JSON ready for analysis
</br></br>**Returns**</br>
listOfMessages : list
</br>&emsp;&emsp;List of messages. Each field in list contains dictionary with message attributes. Structure: `[{}, {}, ..., {}]`.

</br>

`def getReactions(data)`
</br>Retrieve reactions for particular messages in chat.
</br></br>**Parameter**</br>
data : JSON object
</br>&emsp;&emspJSON ready for analysis
</br></br>**Returns**</br>
listOfReactions : list
</br>&emsp;&emspList of reactions with attributes. Each field in list contains dictionary with reactions attributes. Structure: `[{}, {}, ..., {}]`.

</br>

`def getCalls(data)`
</br>Retrieve informations about calls.
</br></br>**Parameter**</br>
data : JSON object
</br>&emsp;&emspJSON ready for analysis
</br></br>**Returns**</br>
listOfCalls : list
</br>&emsp;&emspList of reactions with attributes. Each field in list contains dictionary with calls attributes. Structure: `[{}, {}, ..., {}]`.

</br>

`def getMeetings(data)`
</br>Retrieve informations about meetings in calendar.
</br></br>**Parameter**</br>
data : JSON object
</br>&emsp;&emspJSON ready for analysis
</br></br>**Returns**</br>
listOfMeetings : list
</br>&emsp;&emspList of meetings with attributes. Each field in list contains dictionary with meetings attributes. Structure: `[{}, {}, ..., {}]`.

---
## Printing Methods

`def printConsolePrettyOutput(dictionaryList)`
</br>General print method. 
</br></br>**Parameter**</br>
dictionaryList : list
</br>&emsp;&emspList of dictionaries with data.
</br></br>**Returns**</br>
Print : str
</br>&emsp;&emspPrint output in key : value structure. 

</br>

`def printConsoleMessageThreads(data)`
</br>Special method for print messages in Threads. Responsibilities similar to printConsolePrettyOutput(), but only for messages. 
Threads means separate chats. Every one chat is one thread. 
</br></br>**Parameter**</br>
dictionaryList : list
</br>&emsp;&emspList of dictionaries with data. 
</br></br>**Returns**</br>
Print : str
</br>&emsp;&emspPrint messages in threads.

### Code description

Retrieve list of thread id
```
for thread in range(len(dictionaryList)):  
    conversationId = dictionaryList[thread].get('conversationId')  
    if conversationId not in listOfThreads:  
        listOfThreads.append(conversationId)
```

Create dictionary with messages grouped by conversation id {conversationId: [message, message]}
```
for thread in listOfThreads:  
    tableOfMessages = []  
    counter = 1  
    for row in range(len(dictionaryList)):  
        conversationId = dictionaryList[row].get('conversationId')  
        if thread == conversationId:  
            tableOfMessages.append(dictionaryList[row])  
        counter += 1  
    dictOfMessages[thread] = tableOfMessages
```
Part of printing message emotions.
```
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
```
</br>

`def printHeader(text)`
</br>Method for print header for particular artifacts section. 
</br></br>**Parameter**</br>
text : str
</br>&emsp;&emspText to print.
</br></br>**Returns**</br>
print : str
</br>&emsp;&emspPrint text with some other signs to make it prettier.

---
## Main

Program starting point
### Code description

Program arguments handling.
```
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
```

Input file assignment to variable.
```
inputFile = args.file  
data = fileHandler(inputFile)
```

Handling output file. Additionally checking if name provided by user has .txt extension.
```
outputFile = args.output  
outputFile = outputFile.replace('"', '').replace("'", "")  
if args.output:  
    isExtension = os.path.splitext(args.output)[1]  
    if isExtension.lower() == '.txt':  
        pass  
    else:  
        outputFile += '.txt'
```

Creating output stream with sys.stdout().
```
sys.stdout = open(outputFile, 'a+')  
print(f'Script execution time: {datetime.now(timezone.utc)} UTC')  
print(f'Script running for file {inputFile}')  
print('Script author: Patryk \'Hex7\' Łabuz')  
print(type(data))  
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
```

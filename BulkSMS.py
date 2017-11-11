import requests, logging as log
from openpyxl import Workbook, load_workbook

#Logging Config
log.basicConfig(filename='ActivityLog.log', level=log.DEBUG, format='%(asctime)s %(levelname)-8s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

#Variables
wb = load_workbook('data.xlsx')
ws = wb['Sheet1'] #ws is woksheet [Sheet1] in workbook wb
firstNames = []
lastNames = []
phoneNumbers = []
customFields = []
messages = []
count = 0
countOfMessages = 0

#Data initialization
for column in ws.columns:
    for cell in column:
        if(count is 0):
            firstNames.append(cell.value)
        elif(count is 1):
            lastNames.append(cell.value)
        elif (count is 2):
            phoneNumbers.append(cell.value)
        elif (count is 3):
            customFields.append(cell.value)
        elif (count is 4):
            messages.append(cell.value)
    count+=1


#Functions
def sendText(message, phonenumber):
    url = "https://platform.clickatell.com/messages"
    payload = "{\"content\": \"" + message + "\", \"to\": [\"" + str(phonenumber) + "\"]}"
    headers = {
        'content-type': "application/json",
        'accept': "application/json",
        'authorization': "[auth key]",
        'cache-control': "no-cache",
    }
    response = requests.request("POST", url, data=payload, headers=headers)
    return

def createMessage(firstname, lastname, customfield, message):
    header = firstname+",\\n"
    footer = "\\n[Footer]"
    response = header + message + footer
    return response

def isValid(index):
    if (firstNames[index] != None and phoneNumbers[index] != None and messages[index] != None and lastNames[index] != None and customFields[index] != 0):
        return True
    else:
        return False

for i in range(1,len(firstNames)):
    if(isValid(i)):
        countOfMessages+=1

confirmation = raw_input(str(countOfMessages) + " messages will be sent. Continue? (y/n):")

log.info(str(countOfMessages)+" messages to be sent.")
log.info('User Entered ' + confirmation)

if(confirmation=='y'):
    print "Sending messages...\n"
    for i in range(1, len(firstNames)):
        if (isValid(i)):
            sendText(createMessage(firstNames[i], lastNames[i], customFields[i], messages[i]),phoneNumbers[i])
            log.info(firstNames[i] + ", " + lastNames[i] + ", " + str(phoneNumbers[i]) + " , Sent!")
        else:
            log.warning("Index: " + str(i) + " - Message not sent!")
else:
    print "Program Completed."
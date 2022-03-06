exitStatus = 0
printList = []


def setId(val1, val2, val3):
    if val1:
        global divId, groupId, action, exitStatus
        divId = val1
        groupId = val2
        action = val3
        exitStatus = getExit()
        if exitStatus == 1:
            error = "Client_Error: The process was forced to pause and exit!"
            raise Exception(error)


def getId():
    global divId, groupId, action
    return divId, groupId, action


def setExit(val):
    global exitStatus
    exitStatus = val


def getExit():
    global exitStatus
    return exitStatus


def setListBlank():
    global printList
    printList = []


def setList(val):
    global printList
    printList.append(val)


def getList():
    global printList
    printContent = ""
    for content in printList:
        printContent += '<font color="#D58512">' + str(content) + '</font>' + "<br>"
    return printContent

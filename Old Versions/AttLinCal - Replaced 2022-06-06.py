
# Program Information
program_name = "Sensor Linearity Calibration"
program_author = "Micah Hurd"
program_version = 1.121
python_version = 3.71
cs_number = "CS942153.15"

# Dependent on Installation of Excel Wings (see data to Excel function)
# Use "pip install xlwings"
import scipy.interpolate # Non-native package which must be installed
# import scipy
import re
import math
import statistics
import os
import shutil
import os.path
from os import path
from pathlib import Path
from distutils.dir_util import copy_tree
from tkinter import *
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import time
import tempfile
import logging as log
from columnar import columnar    # Not native

def readConfigFile(filename, searchTag, sFunc="", default_value=''):
    output = "Searched term could not be found"

    try:
        searchTag = searchTag.lower()
        # print("Search Tag: ",searchTag)

        # Open the file
        with open(filename, "r") as filestream:
            # Loop through each line in the file

            for line in filestream:

                if line[0] != "#":

                    currentLine = line
                    equalIndex = currentLine.find('=')
                    if equalIndex != -1:

                        tempLength = len(currentLine)
                        # print("{} {}".format(equalIndex,tempLength))
                        tempIndex = equalIndex
                        configTag = currentLine[0:(equalIndex)]
                        configTag = configTag.lower()
                        configTag = configTag.strip()
                        # print(configTag)

                        configField = currentLine[(equalIndex + 1):]
                        configField = configField.strip()
                        # print(configField)

                        # print("{} {}".format(configTag,searchTag))
                        if configTag == searchTag:

                            # Split each line into separated elements based upon comma delimiter
                            # configField = configField.split(",")

                            # Remove the newline symbol from the list, if present
                            lineLength = len(configField)
                            lastElement = lineLength - 1
                            try:
                                if configField[lastElement] == "\n":
                                    configField.remove("\n")
                            except:
                                configField.strip()
                            # Remove the final comma in the list, if present
                            lineLength = len(configField)
                            lastElement = lineLength - 1

                            try:
                                if configField[lastElement] == ",":
                                    configField = configField[0:lastElement]
                            except:
                                pass

                            lineLength = len(configField)
                            lastElement = lineLength - 1

                            # Apply string manipulation functions, if requested (optional argument)
                            if sFunc != "":
                                sFunc = sFunc.lower()

                                if sFunc == "listout":
                                    configField = configField.split(",")

                                if sFunc == "stringout":
                                    configField = configField.strip("\"")

                                if sFunc == "int":
                                    configField = int(configField)

                                if sFunc == "float":
                                    configField = float(configField)

                            output = configField

    except Exception as e:
        print(f'Configuration Read Error: {e}')
        time.sleep(5)
        output = default_value

    if output == "Searched term could not be found" and default_value != '':
        output = default_value

    filestream.close()
    return output


def writeLog(analysis, logFile):
    import datetime
    import csv
    write_mode = "a"

    currentDT = datetime.datetime.now()

    date_time = currentDT.strftime("%Y-%m-%d %H:%M:%S")

    with open(logFile, mode=write_mode, newline='') as result_file:
        result_writer = csv.writer(result_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

        result_writer.writerow([date_time, analysis])

    result_file.close()

    return 0


def check_and_create_path(path, autocreate=True):
    import os
    error_info = ''
    if not os.path.exists(path):
        if autocreate:
            try:
                os.mkdir(path)
                return True, error_info
            except Exception as error:
                return False, error
        else:
            return False, error_info
    else:
        return True, error_info


def file_check_exists(file):
    import os
    if not os.path.exists(file):
        return False
    else:
        return True


def userInterfaceHeader(program, cs_num, version, cwd, logFile, msg=""):
    print(program + ", Version " + str(version))
    print(f"Procedure: {cs_num}")
    print("Current Working Directory: " + str(cwd))
    print("Log file located at working directory: " + str(logFile))
    print("=======================================================================")
    if msg != "":
        print(msg)
        print("_______________________________________________________________________")
    return 0


def clear():  # Clears the console
    # for windows
    from os import system, name
    if name == 'nt':
        _ = system('cls')

        # for mac and linux(here, os.name is 'posix')
    else:
        _ = system('clear')


def readTxtFile(filename):
    # Place contents of text files into variable
    f = open(filename, 'r')
    x = f.readlines()
    f.close()
    return x


def create_sql_connection(db_file):
    import sqlite3
    from sqlite3 import Error
    """ create a database connection to the SQLite database
        specified by db_file
    :param db_file: database file
    :return: Connection object or None
    """
    #db_file = "C:\\Users\\Micah\\PycharmProjects\\DatabaseTutorial\\standardsDatabase.db"
    print("dbFile", db_file)
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return conn


def select_all_standards(conn):
    import sqlite3
    from sqlite3 import Error
    import functools
    import operator
    """
    Query all rows in the tasks table
    :param conn: the Connection object
    :return:
    """
    recordList = []
    cur = conn.cursor()
    try:
        cur.execute("SELECT * FROM standards")
    except Exception as errorMsg:
        errorMsg = str(errorMsg)
        print(errorMsg)
        if "no such table" in errorMsg:
            create_table(conn)
        standard_info = (666, "Much empty such wow", "------", "1990-01-01", "------", "Delete meeeeeeeeee!")
        create_standard_entry(conn, standard_info, False)
        cur.execute("SELECT * FROM standards")

    rows = cur.fetchall()
    # print(rows)

    for row in rows:
        string = ",".join(map(str, row))
        string = string.split(",")
        # string = "{:15}{:25}{:15}{:15}{:25}{:17}{:30}".format(string[0], string[1], string[2], string[3], string[4],
        #                                                       string[5], string[6])
        recordList.append(string)

    return recordList


def imporStdList(standardDatabaseFile):

    sqlDbConn = create_sql_connection(standardDatabaseFile)

    standardList = select_all_standards(sqlDbConn)

    return standardList


def dBm_to_percent(dBm_initial, dBm_delta):
    dB = dBm_delta - dBm_initial
    percent = ((10 ** (dB / 10)) - 1) * 100
    return percent


def listSelectorGUI(inputList):
    import PySimpleGUI as sg
    import functools
    writeLog("In listSelectorGUI",logFile)
    writeLog("Pulled in standard list: {}".format(inputList), logFile)
    inputList = CheckOverDueStandards(inputList)
    writeLog("Standard List After Filtering Over Due: {}".format(inputList), logFile)

    headerText = "{:15}{:25}{:15}{:15}{:25}{:30}".format("Asset #", "Manufacturer", "Model", "Cal Due", "Cert Description", "Comments")

    outputList = []
    stringOutputList = []
    stringInputList = []
    for index, i in enumerate(inputList):
        listItem = i
        # print("{}\t\t{}\t\t\t{}\t\t\t{}".format(listItem[0],listItem[1],listItem[2],listItem[3]))
        # "{:2}:   {:15}{:30}{:15}{:10}".format(index, listItem[0], listItem[1], listItem[2], listItem[3])
        stringInputList.append("{:15}{:25}{:15}{:15}{:25}{:30}".format(listItem[0], listItem[1], listItem[2], listItem[3], listItem[4],
                                                     listItem[6]))
        #stringInputList.append("{:15}{:30}{:15}{:10}".format(listItem[0], listItem[1], listItem[2], listItem[3]))

    # Populate default standards
    defaultListIndex = 5
    tempIndexList = []
    for index, i in enumerate(inputList):
        listItem = i
        defaultList = listItem[defaultListIndex]
        defaultList = defaultList.lower()

        # defaultFlag = listItem[-1]
        # defaultFlag = defaultFlag.lower()
        # defaultFlag = defaultFlag[0]
        # if defaultFlag == "d":
        if "pscal" in defaultList:
            tempIndexList.append(index)

    if len(tempIndexList) > 0:
        for indexValue in reversed(tempIndexList):
            outputList.append(inputList[indexValue])
            stringOutputList.append(stringInputList[indexValue])
            inputList.pop(indexValue)
            stringInputList.pop(indexValue)

    sg.theme('SystemDefaultForReal')   # Add a little color to your windows
    lineLength = 120
    bt = {'size': (7, 2)}
    lb = {'size': (150,20), 'enable_events': (True), 'font': ('Courier 10')}
    header = {'font': ('Courier 10')}
    layout = [[sg.Text('_' * lineLength)],
              [sg.Text('Available Standards List (Click to add to \"Selected List\")')],
              [sg.Text(headerText,**header)],
              [sg.Listbox(values=stringInputList, key='AvailableStandards',**lb)],
              [sg.Text('_' * lineLength)],
              [sg.Text('Selected Standards List (Click to remove from \"Selected List\")')],
              [sg.Text(headerText,**header)],
              [sg.Listbox(values=stringOutputList, key='SelectedStandards',**lb)],
              [sg.Button('Continue'), sg.Text('(Click Continue once all standards are selected)')]
              ]

    # Create the Window
    window = sg.Window('Select Standards - I built this over the weekend on my own time, just to make your life a little easier; I\'m sure nobody will appreciate it...', layout)
    # Event Loop to process "events"
    while True:
        event, values = window.read()
        # print(event)
        if event in (sg.WIN_CLOSED, 'Continue'):
            break
        if values['AvailableStandards']:  # if something is highlighted in the list
            # sg.popup(f"Your favorite color is {values['AvailableStandards'][0]}")
            index_tuple = window.Element('AvailableStandards').Widget.curselection()
            index = functools.reduce(lambda sub, ele: sub * 10 + ele, index_tuple)
            # print(index)
            outputList.append(inputList[index])
            stringOutputList.append(stringInputList[index])
            inputList.pop(index)
            stringInputList.pop(index)
            # print(outputList)
            window.FindElement('AvailableStandards').Update(values=stringInputList)
            window.FindElement('SelectedStandards').Update(values=stringOutputList)

        if values['SelectedStandards']:
            index_tuple = window.Element('SelectedStandards').Widget.curselection()
            index = functools.reduce(lambda sub, ele: sub * 10 + ele, index_tuple)
            inputList.append(outputList[index])
            stringInputList.append(stringOutputList[index])
            outputList.pop(index)
            stringOutputList.pop(index)
            window.FindElement('AvailableStandards').Update(values=stringInputList)
            window.FindElement('SelectedStandards').Update(values=stringOutputList)

        if event == 'close':
            break

    window.close()

    return outputList


def CheckOverDueStandards(standardsList,tempDebugBool=0):
    from datetime import datetime
    from datetime import date
    import PySimpleGUI as sg2
    writeLog("Inside CheckOverDueStandards",logFile)

    currentDate = datetime.date(datetime.now())
    if tempDebugBool == 1:
        PrintAndLog("Current Date Time: {}".format(currentDate), logFile)

    writeLog("Pulled in standard list: {}".format(standardsList), logFile)
    dueDateIndexLocation = 3
    expiredStandards = []
    for index, i in enumerate(standardsList):
        print(i)
        listItem = i
        # listItem = i.split(",")
        #
        # checkLastItem = listItem[-1]
        # checkLastItem = checkLastItem.lower()
        # checkLastItem = checkLastItem[0]
        #
        # if checkLastItem == "d":
        #     del listItem[-1]


        currentItemDate = listItem[dueDateIndexLocation]
        if tempDebugBool == 1:
            PrintAndLog("currentItemDate: {}".format(currentItemDate), logFile)
        currentItemDateList = currentItemDate.split("-")
        if tempDebugBool == 1:
            PrintAndLog("currentItemDateList: {}".format(currentItemDateList), logFile)
        year = int(currentItemDateList[0])
        month = int(currentItemDateList[1])
        day = int(currentItemDateList[2])

        dateFromList = date(year, month, day)
        if tempDebugBool == 1:
            PrintAndLog("dateFromList: {}".format(dateFromList), logFile)
        if currentDate > dateFromList:
            if tempDebugBool == 1:
                PrintAndLog("currentDate ({}) is > than dateFromList ({})".format(currentDate, dateFromList), logFile)
            expiredStandards.append(i)
            del standardsList[index]
    writeLog("Found these expired standards: {}".format(expiredStandards), logFile)


    if len(expiredStandards) > 0:

        tempString = ''
        for index, i in enumerate(expiredStandards):
            # listItem = i.split(",")
            listItem = i
            # print("{}\t\t{}\t\t\t{}\t\t\t{}".format(listItem[0],listItem[1],listItem[2],listItem[3]))
            tempString += "{:2}:   {:15}{:30}{:15}{:10}\n".format(index, listItem[0], listItem[1], listItem[2], listItem[3])

        sg2.theme('SystemDefaultForReal')  # Add a little color to your windows
        layout = [[sg2.Text('The Following Standards are expired:')],
                  [sg2.Text(tempString, font='Courier 10')],
                  [sg2.Text('These must be updated before they can be used.')],
                  [sg2.Button('Continue')]
                  ]

        # Create the Window
        window = sg2.Window(
            'Expired Standards',
            layout)
        # Event Loop to process "events"
        while True:
            event2, values2 = window.read()
            # print(event)
            if event2 in (sg2.WIN_CLOSED, 'Continue'):
                break
            if event2 == 'close':
                break
        window.close()

    return standardsList


def PrintAndLog(text, logFilePath, printBool=True, logBool=True):
    if printBool:
        print(text)
    if logBool:
        writeLog(text, logFilePath)
    return 0


def UpdateLinearityReferenceDescription(xmlDataList):
    # This function scans the XML data to see if there is linearity data, if so then it
    # Looks for the reference line, and updates the limit "N/A" to "Reference"
    # and it updates the uncertainty field to "N/A"

    writeLog("Attempting to update linearity reference description", logFile)
    FoundLinearity = False  # These should only be false once
    FoundReference = False  # Should only be false once
    AppliedLimitsFix = False
    for index, line in enumerate(xmlDataList):
        lineLowerCase = line.lower()
        LinearitySearchTerm = "linearity"
        ReferenceSearchTerm = "limits"
        ReferenceSearchTerm2 = "n/a"

        if LinearitySearchTerm in lineLowerCase:
            FoundLinearity = True

        if ReferenceSearchTerm in lineLowerCase and ReferenceSearchTerm2 in lineLowerCase:
            FoundReference = True

        if FoundLinearity and FoundReference and (AppliedLimitsFix == False):
            AppliedLimitsFix = True
            writeLog("Found linearity data reference line: {}".format(line), logFile)
            firstWrapper = "<Limits>"
            secondWrapper = '</Limits>'
            value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)
            outputXMLstring = outputXMLstring.replace("val", "N/A")
            newXMLstring = outputXMLstring + '\n'
            xmlDataList[index] = newXMLstring
            writeLog("Wrote new line at index {}: {}".format(index, newXMLstring), logFile)

            tempIndex = index + 1
            line = xmlDataList[tempIndex]
            firstWrapper = "<Uncertainty>"
            secondWrapper = '</Uncertainty>'
            value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)
            outputXMLstring = outputXMLstring.replace("val", "N/A")
            newXMLstring = outputXMLstring + '\n'
            xmlDataList[tempIndex] = newXMLstring
            writeLog("Wrote new line at index{}: {}".format(tempIndex, newXMLstring), logFile)

            tempIndex = index + 2
            line = xmlDataList[tempIndex]
            firstWrapper = "<Pass_Fail>"
            secondWrapper = '</Pass_Fail>'
            value, outputXMLstring = extractValueFromXML(firstWrapper, secondWrapper, line)
            outputXMLstring = outputXMLstring.replace("val", "Reference")
            newXMLstring = outputXMLstring + '\n'
            xmlDataList[tempIndex] = newXMLstring
            writeLog("Wrote new line at index{}: {}".format(tempIndex, newXMLstring), logFile)
            break

    writeLog("Finished looking for linearity reference description", logFile)
    return xmlDataList


def standardize_file_path_format(input_file_path_string):
    import ntpath
    # I wrote this module to standardize file path strings used inside Python programs
    # This allows the differeces between UNIX and Windows style file paths to be eliminated
    # requires the ntpath module
    path, file = ntpath.split(input_file_path_string)
    if path[0] == '\\' and path[1] == '\\':
        path_slash = '\\'
        windows_network_path = True
    elif path[0] == '/' and path[1] == '/':
        path_slash = '\\'
        windows_network_path = True
    else:
        windows_network_path = False
        path_slash = '/'

    if windows_network_path:
        path = path.replace('/', '\\')
    else:
        path = path.replace('\\', '/')

    new_path = ''
    for index, character in enumerate(path):
        last_character = path[index-1]

        if index > 1:
            if character == path_slash and last_character == path_slash:
                character = ''

        new_path += character

    return f'{new_path}{path_slash}{file}'


def printLog(stringMessage, errorInfo=False, console=True):
    # This module allows for printing to the console and to the log in a single line of code; reduces code clutter
    # To use this function you must have imported the logging module and configured it. For example:
    # import logging as log
    # log.basicConfig(filename=logFile, filemode='a', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S', level=logging.INFO)

    if console:
        print(stringMessage)

    if errorInfo:
        log.info(stringMessage, exc_info=errorInfo)
    else:
        log.info(stringMessage)


def rename_if_file_exists(filename):
    import os
    import pathlib
    new_filename = filename

    file_exists = os.path.exists(filename)

    cntr = 0
    while file_exists:
        cntr+=1
        file_extension = pathlib.Path(new_filename).suffix
        filename_wout_extension = filename.strip(file_extension)

        copy_str = " - Copy("
        if copy_str in filename_wout_extension:
            index_loc = filename_wout_extension.find(copy_str)
            filename_wout_extension = filename_wout_extension[:index_loc]

        new_filename = f'{filename_wout_extension} - Copy({cntr}){file_extension}'

        file_exists = os.path.exists(new_filename)

    return new_filename


def dBtoPercent(dB_input):
    return ((10 ** ((dB_input) / 10)) - 1) * 100


def check_unc_budget(budget_file_path):
    printLog('Checking linearity budget file...')
    budget_expired = True
    while budget_expired:
        budget_expired = UncertaintyBudget.check_is_expired(budget_file_path)
        if budget_expired:
            temp_msg = f'The linearity budget file referenced below is expired! Please update before proceeding.\n\n' \
                       f'{budget_file_path}'
            msg_box_simple(temp_msg)

    # input(f"Budget check: {budget_expired}")


def writeListToFile(filename, my_list, write_type='w'):
    def ensure_file_exists(filepath):
        if not path.exists(filepath):
            with open(filepath, 'w') as fp:
                pass

    def write_list_normal(filename, write_type):
        with open(filename, write_type) as f:
            for item in my_list:
                f.write("%s\n" % item)

    def write_list_utf8(filename, write_type):
        with open(filename, write_type, encoding="utf-8") as f:
            for item in my_list:
                f.write("%s\n" % item)

    ensure_file_exists(filename)

    try:
        write_list_normal(filename, write_type)
    except Exception as e:
        error = f'{e}'
        error = error.lower()
        if 'permission denied' in error:
            temp_str = f'Access Denied for file: {filename}\n\n' \
                       f'Please ensure the file is not in use by another program before proceeding!'
            msg_box_simple(temp_str)
            write_list_normal(filename, write_type)
        else:
            write_list_utf8(filename, write_type)


def get_convert_timestamp(ts='', format=None):
    from datetime import datetime

    if ts == "":
        ct = datetime.now()
        ts = ct.timestamp()
    else:
        ts = float(ts)


    dt_object = datetime.fromtimestamp(ts)

    if format == 'time':
        date_time = dt_object.strftime("%H:%M:%S")
        return date_time
    elif format == 'date_time':
        date_time = dt_object.strftime("%Y-%m-%d_%H:%M:%S")
    else:
        date_time = dt_object.strftime("%Y-%m-%d")

    return date_time


def setSigDigits(value, qtySigDigitsRequired):
    import math

    valueAsFloat = float(value)

    if qtySigDigitsRequired < 1:
        qtySigDigitsRequired = 1

    SciNotationDigits = qtySigDigitsRequired - 1

    NewValue = "{:0.{}e}".format(valueAsFloat, SciNotationDigits)

    index = NewValue.find("e")
    qtyOfZeros = NewValue[(index + 1):]
    qtyOfZeros = int(qtyOfZeros)

    if qtyOfZeros < 0:
        qtyOfZeros = abs(qtyOfZeros) - 1
        BaseNumber = str(NewValue[:index])
        BaseNumber = BaseNumber.replace(".", "")

        NewStringValue = "0."
        for i in range(qtyOfZeros):
            NewStringValue += "0"
        NewStringValue += BaseNumber
    elif qtyOfZeros == 0:
        NewStringValue = NewValue[:index]
    elif qtyOfZeros > 0:
        NewStringValue = float(NewValue)
        CheckString = str(NewStringValue)
        FloorValue = math.floor(NewStringValue)

        remainder = NewStringValue - FloorValue
        if remainder == 0:
            NewStringValue = int(NewStringValue)

    return str(NewStringValue)


def checkUncBudget(budgetTxtFile,uncVal,uncFreq,uncMsmt=0.0):
    import datetime
    from datetime import date

    # Suppress scientific notation (this code ended up being unnecessary)
    # uncVal = f'{uncVal:.20f}'                       # 20 digits, because why not?
    # uncFreq = f'{uncFreq:.0f}'                      # Frequency in Hz should have no resolution after the decimal

    # Place contents of XML files into variable
    f = open(budgetTxtFile, 'r')
    budgetData = f.readlines()
    f.close()

    # Strip any existing whitespace out of the list elements
    for index, element in enumerate(budgetData):
        if element == "\n":
            del budgetData[index]
        else:
            newElement = element.strip()
            budgetData[index] = newElement



    # Placed current date into variable
    today = date.today()
    currentDate = today.strftime("%Y-%m-%d")

    # Obtain the line containing the expiration date of the budget file
    tempList = budgetData[0]
    tempList = tempList.strip()
    tempList = tempList.split(",")
    fileDate = tempList[1]
    fileDate = fileDate.split("-")                  # Store date elements into a list
    currentDate = currentDate.split("-")            # Also store today's date elements into list

    # Convert dates into formats which can be compared against
    fileDateInFormat = datetime.datetime(int(fileDate[0]), int(fileDate[1]), int(fileDate[2]))
    currentDateInFormat = datetime.datetime(int(currentDate[0]), int(currentDate[1]), int(currentDate[2]))

    # Check to see if the file date has been exceeded
    if currentDateInFormat > fileDateInFormat:
        return "File > {} < expiration date is exceeded!".format(budgetTxtFile)

    del budgetData[0]                            # Delete the list element containing the date

    # Check to see if this budget format includes Power ranges
    tempList = budgetData[0].split(",")
    tempQty = len(tempList)
    if tempQty > 2:
        pRangePresent = True
    else:
        pRangePresent = False

    # Break out the elements of the uncertainty budget list
    freqList = []
    rangeList = []
    uncList = []
    for index, element in enumerate(budgetData):
        tempList = element.split(",")
        if pRangePresent == True:
            freqList.append(tempList[0])
            rangeList.append(tempList[1])
            uncList.append(float(tempList[2]))
        else:
            freqList.append(tempList[0])
            uncList.append(float(tempList[1]))

    # print(freqList)
    # print(rangeList)
    # print(uncList)

    # Find the current uncertainty frequency within the budget frequency list
    for index, element in enumerate(freqList):
        tempList = element.split(">")
        startFreq = float(tempList[0])
        stopFreq = float(tempList[1])
        uncFreq = float(uncFreq)


        if (uncFreq >= startFreq) and (uncFreq <= stopFreq):

            if pRangePresent == True:
                tempList = rangeList[index].split(">")
                startPow = float(tempList[0])
                stopPow = float(tempList[1])

                if (uncMsmt >= startPow) and (uncMsmt <= stopPow):
                    tempUncListVal = uncList[index]

                    if uncVal < tempUncListVal:
                        uncVal = tempUncListVal
                        return uncVal
            else:
                tempUncListVal = uncList[index]
                if uncVal < tempUncListVal:
                    uncVal = tempUncListVal
                    return uncVal


    return uncVal


def dBm_mW(dBm):
    value_mW = 10**(dBm/10)
    return value_mW


def mW_dBm(mW):
    import math

    value_dBm = 10 * math.log10(mW)
    return value_dBm


def percent_to_dB(value_as_a_percent):
    import math
    dB = 10 * math.log10((value_as_a_percent + 100) / 100)
    return dB


def Students_T_Lookup(DegreesOfFreedom, Confidence=95.45):
    # This function pulls the student's T coverage factor from a series of lists, corresponding to the user input
    # degrees of freedom and a selectable confidence interval of 90%, 95%, 95.45%, 99%, 99.5%, 99.73%, and 99.9%
    # The function defaults to 95.45% if no confidence interval is entered
    # If degrees of freedom is set to zero then the function assigns infinite degrees of freedom

    Confidence_90 = [1.645, 6.314, 2.920, 2.353, 2.132, 2.015, 1.943, 1.895, 1.860, 1.833, 1.812, 1.796, 1.782, 1.771,
                     1.761, 1.753, 1.746, 1.740, 1.734, 1.729, 1.725, 1.721, 1.717, 1.714, 1.711, 1.708, 1.706, 1.703,
                     1.701, 1.699, 1.697, 1.696, 1.694, 1.692, 1.691, 1.690, 1.688, 1.687, 1.686, 1.685, 1.684, 1.683,
                     1.682, 1.681, 1.680, 1.679, 1.679, 1.678, 1.677, 1.677, 1.676, 1.675, 1.675, 1.674, 1.674, 1.673,
                     1.673, 1.672, 1.672, 1.671, 1.671, 1.670, 1.670, 1.669, 1.669, 1.669, 1.668, 1.668, 1.668, 1.667,
                     1.667, 1.667, 1.666, 1.666, 1.666, 1.665, 1.665, 1.665, 1.665, 1.664, 1.664, 1.664, 1.664, 1.663,
                     1.663, 1.663, 1.663, 1.663, 1.662, 1.662, 1.662, 1.662, 1.662, 1.661, 1.661, 1.661, 1.661, 1.661,
                     1.661, 1.660, 1.660]
    Confidence_95 = [1.960, 12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228, 2.201, 2.179, 2.160,
                     2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086, 2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052,
                     2.048, 2.045, 2.042, 2.040, 2.037, 2.035, 2.032, 2.030, 2.028, 2.026, 2.024, 2.023, 2.021, 2.020,
                     2.018, 2.017, 2.015, 2.014, 2.013, 2.012, 2.011, 2.010, 2.009, 2.008, 2.007, 2.006, 2.005, 2.004,
                     2.003, 2.002, 2.002, 2.001, 2.000, 2.000, 1.999, 1.998, 1.998, 1.997, 1.997, 1.996, 1.995, 1.995,
                     1.994, 1.994, 1.993, 1.993, 1.993, 1.992, 1.992, 1.991, 1.991, 1.990, 1.990, 1.990, 1.989, 1.989,
                     1.989, 1.988, 1.988, 1.988, 1.987, 1.987, 1.987, 1.986, 1.986, 1.986, 1.986, 1.985, 1.985, 1.985,
                     1.984, 1.984, 1.984]
    Confidence_9545 = [2.000, 13.968, 4.527, 3.307, 2.869, 2.649, 2.517, 2.429, 2.366, 2.320, 2.284, 2.255, 2.231,
                       2.212, 2.195, 2.181, 2.169, 2.158, 2.149, 2.140, 2.133, 2.126, 2.120, 2.115, 2.110, 2.105, 2.101,
                       2.097, 2.093, 2.090, 2.087, 2.084, 2.081, 2.079, 2.076, 2.074, 2.072, 2.070, 2.068, 2.066, 2.064,
                       2.063, 2.061, 2.060, 2.058, 2.057, 2.056, 2.055, 2.053, 2.052, 2.051, 2.050, 2.049, 2.048, 2.047,
                       2.046, 2.046, 2.045, 2.044, 2.043, 2.043, 2.042, 2.041, 2.040, 2.040, 2.039, 2.039, 2.038, 2.037,
                       2.037, 2.036, 2.036, 2.035, 2.035, 2.034, 2.034, 2.033, 2.033, 2.033, 2.032, 2.032, 2.031, 2.031,
                       2.031, 2.030, 2.030, 2.029, 2.029, 2.029, 2.028, 2.028, 2.028, 2.028, 2.027, 2.027, 2.027, 2.026,
                       2.026, 2.026, 2.026, 2.025]
    Confidence_99 = [2.576, 63.657, 9.925, 5.841, 4.604, 4.032, 3.707, 3.499, 3.355, 3.250, 3.169, 3.106, 3.055, 3.012,
                     2.977, 2.947, 2.921, 2.898, 2.878, 2.861, 2.845, 2.831, 2.819, 2.807, 2.797, 2.787, 2.779, 2.771,
                     2.763, 2.756, 2.750, 2.744, 2.738, 2.733, 2.728, 2.724, 2.719, 2.715, 2.712, 2.708, 2.704, 2.701,
                     2.698, 2.695, 2.692, 2.690, 2.687, 2.685, 2.682, 2.680, 2.678, 2.676, 2.674, 2.672, 2.670, 2.668,
                     2.667, 2.665, 2.663, 2.662, 2.660, 2.659, 2.657, 2.656, 2.655, 2.654, 2.652, 2.651, 2.650, 2.649,
                     2.648, 2.647, 2.646, 2.645, 2.644, 2.643, 2.642, 2.641, 2.640, 2.640, 2.639, 2.638, 2.637, 2.636,
                     2.636, 2.635, 2.634, 2.634, 2.633, 2.632, 2.632, 2.631, 2.630, 2.630, 2.629, 2.629, 2.628, 2.627,
                     2.627, 2.626, 2.626]
    Confidence_995 = [2.807, 127.321, 14.089, 7.453, 5.598, 4.773, 4.317, 4.029, 3.833, 3.690, 3.581, 3.497, 3.428,
                      3.372, 3.326, 3.286, 3.252, 3.222, 3.197, 3.174, 3.153, 3.135, 3.119, 3.104, 3.091, 3.078, 3.067,
                      3.057, 3.047, 3.038, 3.030, 3.022, 3.015, 3.008, 3.002, 2.996, 2.990, 2.985, 2.980, 2.976, 2.971,
                      2.967, 2.963, 2.959, 2.956, 2.952, 2.949, 2.946, 2.943, 2.940, 2.937, 2.934, 2.932, 2.929, 2.927,
                      2.925, 2.923, 2.920, 2.918, 2.916, 2.915, 2.913, 2.911, 2.909, 2.908, 2.906, 2.904, 2.903, 2.902,
                      2.900, 2.899, 2.897, 2.896, 2.895, 2.894, 2.892, 2.891, 2.890, 2.889, 2.888, 2.887, 2.886, 2.885,
                      2.884, 2.883, 2.882, 2.881, 2.880, 2.880, 2.879, 2.878, 2.877, 2.876, 2.876, 2.875, 2.874, 2.873,
                      2.873, 2.872, 2.871, 2.871]
    Confidence_9973 = [3.000, 235.784, 19.206, 9.219, 6.620, 5.507, 4.904, 4.530, 4.277, 4.094, 3.957, 3.850, 3.764,
                       3.694, 3.636, 3.586, 3.544, 3.507, 3.475, 3.447, 3.422, 3.400, 3.380, 3.361, 3.345, 3.330, 3.316,
                       3.303, 3.291, 3.280, 3.270, 3.261, 3.252, 3.244, 3.236, 3.229, 3.222, 3.216, 3.210, 3.204, 3.199,
                       3.194, 3.189, 3.184, 3.180, 3.175, 3.171, 3.168, 3.164, 3.160, 3.157, 3.154, 3.151, 3.148, 3.145,
                       3.142, 3.140, 3.137, 3.135, 3.132, 3.130, 3.128, 3.126, 3.123, 3.121, 3.120, 3.118, 3.116, 3.114,
                       3.112, 3.111, 3.109, 3.108, 3.106, 3.105, 3.103, 3.102, 3.100, 3.099, 3.098, 3.096, 3.095, 3.094,
                       3.093, 3.092, 3.091, 3.090, 3.089, 3.087, 3.086, 3.085, 3.085, 3.084, 3.083, 3.082, 3.081, 3.080,
                       3.079, 3.078, 3.078, 3.077]
    Confidence_999 = [3.291, 636.619, 31.599, 12.924, 8.610, 6.869, 5.959, 5.408, 5.041, 4.781, 4.587, 4.437, 4.318,
                      4.221, 4.140, 4.073, 4.015, 3.965, 3.922, 3.883, 3.850, 3.819, 3.792, 3.768, 3.745, 3.725, 3.707,
                      3.690, 3.674, 3.659, 3.646, 3.633, 3.622, 3.611, 3.601, 3.591, 3.582, 3.574, 3.566, 3.558, 3.551,
                      3.544, 3.538, 3.532, 3.526, 3.520, 3.515, 3.510, 3.505, 3.500, 3.496, 3.492, 3.488, 3.484, 3.480,
                      3.476, 3.473, 3.470, 3.466, 3.463, 3.460, 3.457, 3.454, 3.452, 3.449, 3.447, 3.444, 3.442, 3.439,
                      3.437, 3.435, 3.433, 3.431, 3.429, 3.427, 3.425, 3.423, 3.421, 3.420, 3.418, 3.416, 3.415, 3.413,
                      3.412, 3.410, 3.409, 3.407, 3.406, 3.405, 3.403, 3.402, 3.401, 3.399, 3.398, 3.397, 3.396, 3.395,
                      3.394, 3.393, 3.392, 3.390]

    if DegreesOfFreedom > 100:
        DegreesOfFreedom = 100
    elif DegreesOfFreedom < 0:
        DegreesOfFreedom = 1

    if Confidence == 90:
        CoverageFactor = Confidence_90[DegreesOfFreedom]
    elif Confidence == 95:
        CoverageFactor = Confidence_95[DegreesOfFreedom]
    elif Confidence == 99:
        CoverageFactor = Confidence_99[DegreesOfFreedom]
    elif Confidence == 99.5:
        CoverageFactor = Confidence_995[DegreesOfFreedom]
    elif Confidence == 99.73:
        CoverageFactor = Confidence_9973[DegreesOfFreedom]
    elif Confidence == 99.9:
        CoverageFactor = Confidence_999[DegreesOfFreedom]
    else:
        CoverageFactor = Confidence_9545[DegreesOfFreedom]

    return CoverageFactor


def printInLine(inputString):
    import os
    maxLength = 121
    os.system('mode con: cols={} lines=40'.format(maxLength+1))
    inputString = str(inputString)
    inputStringLength = len(inputString)
    stringAdder = maxLength - inputStringLength
    if stringAdder < 0:
        inputString = inputString[:stringAdder]
        stringAdder = 0

    inputString += (" " * stringAdder)
    print("{}\r\r".format(inputString), end="")


def queryVisa(instrument, command, sFunc="", retryQty=5):
    command = str(command)

    i = 0
    while i < retryQty:
        try:
            msmt = (instrument.query(command))
            i = 6
        except:
            msmt = "No Data Received"
            time.sleep(0.25)
            i += 1

    if msmt.startswith('\"') and msmt.endswith('\"'):
        msmt = msmt[1:-1]
    # msmt = float(msmt)
    msmt = msmt.strip()

    if sFunc != "":
        sFunc = sFunc.lower()

        if sFunc == "float":
            msmt = float(msmt)

    return msmt


def writeVisa(instrument, command, opc=False, opc_cmd="*OPC?", response="+1", timeout=10, msg="Waiting for instrument to complete the operation..."):
    e = 'No Error Received.'
    inst = instrument
    retry = True
    cnt = 0
    while retry:
        cnt+=1
        if cnt == 3:
            retry = False
        try:
            inst.write(command)
            retry = False
        except Exception as e:
            temp_msg = f'Check connections and address for resource:\n' \
                       f'{inst}\n\n' \
                       f'Received error: \n' \
                       f'{e}'

            msg_box_simple(temp_msg)
        if cnt == 3:
            temp_msg = f'Could not communicate with resource:\n' \
                       f'{inst}\n\n' \
                       f'Error: \n{e}'
            error_and_exit(temp_msg)


    temp_msg = f'Command Sent: {command}'

    if opc:
        inst_response = visa_OPC_handler(instrument, cmd=opc_cmd, response=response, timeout=timeout,
                         msg=msg)
        temp_msg = f'Command Sent: {command}, Response: {inst_response}'

    return temp_msg


def visa_OPC_handler(visaResourceString, cmd="*OPC?", response="+1", timeout=10, msg="Waiting for instrument to complete the operation..."):
    tempBool = False
    firstRunBool = False
    visaRX = ""
    counter = 0
    while tempBool == False:
        if firstRunBool == True:
            time.sleep(1)
            counter += 1
            printInLine("{}, completion request number {}".format(msg, counter))

        visaRX = queryVisa(visaResourceString, cmd)
        firstRunBool = True

        if (visaRX in response) or (response in visaRX):
            tempBool = True
            return visaRX

        if counter >= timeout:
            input("The Visa {} request has timed out; press enter to continue".format(cmd))
            tempBool = True
            return -1
    print("\n\n")
    clear()


def get_config_file_settings():
    try:
        global debug
        global debug_flag
        debug = readConfigFile(configFile, "debug", "int")
        printLog("Debug on or off: {}.".format(debug))
        try:
            debug_flag = True if debug == 1 else False
        except Exception as err:
            temp_msg = f'Attemtped to set debug_flag boolean; got error:\n{err}'
            printLog(temp_msg)
            error_and_exit(temp_msg)
        printLog("Debug flag set to {}.".format(debug_flag))

        global verbose
        global verbose_flag
        verbose = readConfigFile(configFile, "verbose", "int")
        printLog("verbose on or off: {}.".format(verbose))
        try:
            verbose_flag = True if verbose == 1 else False
        except Exception as err:
            temp_msg = f'Attemtped to set verbose_flag boolean; got error:\n{err}'
            printLog(temp_msg)
            error_and_exit(temp_msg)
        printLog("verbose flag set to {}.".format(verbose_flag))

        global msmt_templates_folder
        msmt_templates_folder = readConfigFile(configFile, "msmt_templates_folder")
        printLog("msmt_templates_folder: {}.".format(msmt_templates_folder))

        global msmt_results_folder
        msmt_results_folder = readConfigFile(configFile, "msmt_results_folder")
        printLog("msmt_results_folder: {}.".format(msmt_results_folder))

        global PS_CalResultsFolder
        PS_CalResultsFolder = readConfigFile(configFile, "PS_CalResultsFolder")
        printLog("PS_CalResultsFolder: {}.".format(PS_CalResultsFolder))

        global exercise_att
        exercise_att = readConfigFile(configFile, "exercise_att", "int")
        printLog("exercise_att on or off: {}.".format(exercise_att))
        try:
            exercise_att = True if exercise_att == 1 else False
        except Exception as err:
            temp_msg = f'Attemtped to set exercise_att boolean; got error:\n{err}'
            exercise_att = True
        printLog("exercise_att set to {}.".format(debug_flag))

        global standardsDataFolder
        standardsDataFolder = readConfigFile(configFile, "standardsDataFolder")
        printLog("standardsDataFolder: {}.".format(standardsDataFolder))

        global numberSigDigits
        numberSigDigits = readConfigFile(configFile, "numberSigDigits", "int")
        printLog("Number of significant digits to correct to: {}.".format(numberSigDigits))

        global linBudgetTxtFile
        linBudgetTxtFile = readConfigFile(configFile, "linBudgetTxtFile")
        printLog("Location of linearity budget: {}.".format(linBudgetTxtFile))

        global linearityCalDataFilePath11
        linearityCalDataFilePath11 = readConfigFile(configFile, "linearityCalDataFilePath11")
        printLog("File location of the 11 dB Step Attenuator Cal Data: {}.".format(linearityCalDataFilePath11))

        global linearityCalDataFilePath110
        linearityCalDataFilePath110 = readConfigFile(configFile, "linearityCalDataFilePath110")
        printLog("File location of the 110 dB Step Attenuator Cal Data: {}.".format(linearityCalDataFilePath110))

        global generator_driver
        generator_driver = readConfigFile(configFile, "generator_driver")
        printLog("File location of the Signal Generator Driver File: {}.".format(generator_driver))

        global pm_driver
        pm_driver = readConfigFile(configFile, "pm_driver")
        printLog("File location of the Power Meter Driver File: {}.".format(pm_driver))

        global attenuator_driver
        attenuator_driver = readConfigFile(configFile, "attenuator_driver")
        printLog("File location of the Electronic Step Attenuator Driver File: {}.".format(attenuator_driver))

        global plot_x_inches
        plot_x_inches = readConfigFile(configFile, "plot_x_inches", "int")
        printLog("plot_x_inches: {}.".format(plot_x_inches))

        global plot_y_inches
        plot_y_inches = readConfigFile(configFile, "plot_y_inches", "int")
        printLog("plot_y_inches: {}.".format(plot_y_inches))

        global normalize
        global normalize_flag
        normalize = readConfigFile(configFile, "normalize", "int")
        printLog("normalize on or off: {}.".format(normalize))
        try:
            normalize_flag = True if normalize == 1 else False
        except Exception as err:
            temp_msg = f'Attemtped to set normalize_flag boolean; got error:\n{err}'
            printLog(temp_msg)
            error_and_exit(temp_msg)
        printLog("normalize_flag set to {}.".format(normalize_flag))


    except:
        printLog("Exception occurred while pulling configuration values", errorInfo=True)
        exit()


def verify_file_paths(file_check_list):
    overall_temp_bool = True
    false_list = []
    for item in file_check_list:
        tempBool = file_check_exists(item)

        if tempBool == False:
            overall_temp_bool = False
            printLog(f'File or Folder path from configuration file does not exist: {item}')
            false_list.append(item)

    if overall_temp_bool == False:

        temp_str = f'One or more configuration File or Folder path did not exist: \n\n'
        for item in false_list:
            temp_str = temp_str + f'- {item}\n'

        temp_str = temp_str + f'\nPlease fix the broken files / paths and re-run the program.'
        printLog(temp_str)
        error_and_exit(temp_str)


def check_lin_lists(step_list, tol_list):
    new_steps_list = []
    new_tol_list = []

    if step_list[0] > step_list[1]:

        for item in step_list[::-1]:
            new_steps_list.append(item)

        for item in tol_list[::-1]:
            new_tol_list.append(item)
    else:
        new_steps_list = step_list
        new_tol_list = tol_list
    return (new_steps_list, new_tol_list)


def get_dut_template_data(template_filepath=''):
    if not template_filepath == '':
        templateFile = template_filepath

    # if debug_flag == False:
    #     templateFile = template_filepath
    # else:
    #     templateFile = 'C:\\Users\\Micah\\PycharmProjects\\AttLinCal\\U848xA.lin'

    # Get required variables from the DUT Measurement template file
    global dutModel
    dutModel = readConfigFile(templateFile, "dutModel")
    printLog(f"Template File Model: {dutModel}")

    if debug_flag == False:
        tempString = f'The selected measurement template file is for model:\n\n{dutModel}\n\nIs this correct?'
        if not yes_no_popup_simple(tempString):
            tempString = f'User selected template for model {dutModel} which is incorrect. Please choose the ' \
                         f'correct measurement template file.'
            msg_box_simple(tempString)
            return False

    print("Importing parameters from the " + str(dutModel) + " measurement template...")
    printLog(f"Importing parameters from measurement template file: {templateFile}")
    time.sleep(0.5)


    global generator_name
    generator_name = readConfigFile(templateFile, "sGen")
    printLog(f"generator_name: {generator_name}", console=False)

    global sGenVisaResource
    try:
        sGenVisaResource = readConfigFile(templateFile, "sGenVisaResource")
    except:
        sGenVisaResource = ""
    printLog(f"sGenVisaResource: {sGenVisaResource}", console=False)

    global pMeterName
    pMeterName = readConfigFile(templateFile, "pMeter")
    printLog(f"pMeterName: {pMeterName}", console=False)

    global pMeterVisaResourceIdent
    try:
        pMeterVisaResourceIdent = readConfigFile(templateFile, "pMeterVisaResource")
    except:
        pMeterVisaResourceIdent = ""
    printLog(f"pMeterVisaResourceIdent: {pMeterVisaResourceIdent}", console=False)

    global step_att_name
    step_att_name = readConfigFile(templateFile, "stepAttenuator")
    printLog(f"step_att_name: {step_att_name}", console=False)

    global stepAttVisaResourceIdent
    try:
        stepAttVisaResourceIdent = readConfigFile(templateFile, "stepAttVisaResource")
    except:
        stepAttVisaResourceIdent = ""
    printLog(f"stepAttVisaResourceIdent: {stepAttVisaResourceIdent}", console=False)

    global uom
    uom = readConfigFile(templateFile, "uom")
    uom = uom.lower()
    printLog(f"uom: {uom}", console=False)
    if uom != 'db' and uom != 'searched term could not be found':
        error_and_exit(f'Unit of measure from the DUT template file equals: {uom}. UOM \"dB\" supported only!')

    global biasMsmtQty
    temp_value = readConfigFile(templateFile, "biasMsmtQty")
    biasMsmtQty = sanitize_variable(temp_value, default_response=30, specified_class='int',
                                    eval_operation='at least',
                                    eval_threshold=1)
    printLog(f"biasMsmtQty - Template File Value: {temp_value}, Sanitized Value: {biasMsmtQty}", console=False)

    global settlingTime
    temp_value = readConfigFile(templateFile, "settlingTime")
    settlingTime = sanitize_variable(temp_value, default_response=5, specified_class='int',
                                    eval_operation='at least',
                                    eval_threshold=1)
    printLog(f"settlingTime - Template File Value: {temp_value}, Sanitized Value: {settlingTime}", console=False)

    global samplingQuantity
    temp_value = readConfigFile(templateFile, "samplingQuantity", sFunc="int")
    samplingQuantity = sanitize_variable(temp_value, default_response=3, specified_class='int',
                                     eval_operation='at least',
                                     eval_threshold=1)
    printLog(f"samplingQuantity - Template File Value: {temp_value}, Sanitized Value: {samplingQuantity}", console=False)

    global sampling_intv
    temp_value = readConfigFile(templateFile, "sampling_intv", sFunc="float")
    sampling_intv = sanitize_variable(temp_value, default_response=1, specified_class='float',
                                         eval_operation='at least',
                                         eval_threshold=0.001)
    printLog(f"sampling_intv - Template File Value: {temp_value}, Sanitized Value: {sampling_intv}", console=False)



    global test_frequency
    test_frequency = readConfigFile(templateFile, "test_frequency")
    printLog(f"test_frequency: {test_frequency}", console=False)

    global excelSource
    excelSource = readConfigFile(templateFile, "excelSource")
    printLog(f"excelSource: {excelSource}", console=False)

    global rowOffset
    rowOffset = readConfigFile(templateFile, "rowOffset")
    printLog(f"rowOffset: {rowOffset}", console=False)

    global msmtFileSheet
    msmtFileSheet = readConfigFile(templateFile, "excelSheetName")
    printLog(f"msmtFileSheet: {msmtFileSheet}", console=False)

    global pdfMerge
    pdfMerge = readConfigFile(templateFile, "pdfMerge")
    pdfMerge = pdfMerge.lower()
    printLog(f"pdfMerge: {pdfMerge}", console=False)

    global lin_steps_list
    lin_steps_list = readConfigFile(templateFile, "linSteps", sFunc='listout')
    for index, item in enumerate(lin_steps_list):
        stripped = item.strip()
        stripped = float(stripped)
        lin_steps_list[index] = stripped
    printLog(f"lin_steps_list: {lin_steps_list}", console=False)

    checked_good = False
    while not checked_good:
        global lin_steps_tol
        lin_steps_tol = readConfigFile(templateFile, "tol", sFunc='listout')
        for index, item in enumerate(lin_steps_tol):
            stripped = item.strip()
            stripped = float(stripped)
            lin_steps_tol[index] = stripped
        printLog(f"lin_steps_tol: {lin_steps_tol}", console=False)

        # Ensure the lin steps list, and tol list, from the config file are in the correct order
        try:
            lin_steps_list, lin_steps_tol = check_lin_lists(lin_steps_list, lin_steps_tol)
            checked_good = True
        except Exception as err:

            temp_msg = f'Failed to confirm / re-order the linearity step and tolerence lists:\n' \
                       f'{err}'
            error_and_exit(temp_msg)

        if not len(lin_steps_list) == len(lin_steps_tol):
            checked_good = False
            temp_msg = f'The linearity steps list is a different length than the tolerance list.\n' \
                       f'These lists must be equal length in the DUT template file:\n' \
                       f'{templateFile}\n\n' \
                       f'Open the template file ande correct this before proceeding'
            msg_box_simple(temp_msg)
        else:
            checked_good = True


    global refStepSetting
    refStepSetting = readConfigFile(templateFile, "refStepSetting")
    printLog(f"refStepSetting: {refStepSetting}", console=False)
    try:
        refStepSetting = int(refStepSetting)
    except Exception as err:
        temp_msg = f'Attempted to cast referencePower from config file as an integer; failed:\n {err}\n\n' \
                   f'referencePower must be an integer!'
        printLog(temp_msg)
        error_and_exit(temp_msg)

    # Ensure the reference setting value is with the attenuation step list
    if not refStepSetting in lin_steps_list:
        temp_msg = f'Reference step value\n\n {refStepSetting} dB\n\n' \
                   f'is not within the linearity steps list:\n\n' \
                   f'{lin_steps_list} dB steps\n\n' \
                   f'Please add the desired reference setting to the linearity step list.'
        error_and_exit(temp_msg)

# Get Signal Generator Remote Commands
    global sGenID
    sGenID = readConfigFile(generator_driver, "sGenID")
    printLog(f"sGenID: {sGenID}", console=False)

    global sGenRst
    sGenRst = readConfigFile(generator_driver, "sGenRst")
    printLog(f"sGenRst: {sGenRst}", console=False)

    global sGenFreqSet
    sGenFreqSet = readConfigFile(generator_driver, "sGenFreqSet")
    printLog(f"sGenFreqSet: {sGenFreqSet}", console=False)

    global sGenPowSet
    sGenPowSet = readConfigFile(generator_driver, "sGenPowSet")
    printLog(f"sGenPowSet: {sGenPowSet}", console=False)

    global sGenFreqRead
    sGenFreqRead = readConfigFile(generator_driver, "sGenFreqRead")
    printLog(f"sGenFreqRead: {sGenFreqRead}", console=False)

    global sGenPowRead
    sGenPowRead = readConfigFile(generator_driver, "sGenPowRead")
    printLog(f"sGenPowRead: {sGenPowRead}", console=False)

    global sGenConfig
    sGenConfig = readConfigFile(generator_driver, "sGenConfig")
    printLog(f"sGenConfig: {sGenConfig}", console=False)

    global sGenOn
    sGenOn = readConfigFile(generator_driver, "sGenOn")
    printLog(f"sGenOn: {sGenOn}", console=False)

    global sGenOff
    sGenOff = readConfigFile(generator_driver, "sGenOff")
    printLog(f"sGenOff: {sGenOff}", console=False)

    global unlevel_err_check
    unlevel_err_check = readConfigFile(generator_driver, "unlevel_err_check")
    printLog(f"unlevel_err_check: {unlevel_err_check}", console=False)

    global unlevel_err_response
    unlevel_err_response = readConfigFile(generator_driver, "unlevel_err_response")
    printLog(f"unlevel_err_response: {unlevel_err_response}", console=False)


    # Get Power Meter Driver Commands
    global pmID
    pmID = readConfigFile(pm_driver, "pmID")
    printLog(f"pmID: {pmID}", console=False)

    global pmRst
    pmRst = readConfigFile(pm_driver, "pmRst")
    printLog(f"pmRst: {pmRst}", console=False)

    global pmZero
    pmZero = readConfigFile(pm_driver, "pmZero")
    printLog(f"pmZero: {pmZero}", console=False)

    global pmCal
    pmCal = readConfigFile(pm_driver, "pmCal")
    printLog(f"pmCal: {pmCal}", console=False)

    global pmFreq
    pmFreq = readConfigFile(pm_driver, "pmFreq")
    printLog(f"pmFreq: {pmFreq}", console=False)

    global pmdBmeas
    pmdBmeas = readConfigFile(pm_driver, "pmdBmeas")
    printLog(f"pmdBmeas: {pmdBmeas}", console=False)

    global pmConfig
    pmConfig = readConfigFile(pm_driver, "pmConfig")
    printLog(f"pmConfig: {pmConfig}", console=False)

    global pmOpc
    pmOpc = readConfigFile(pm_driver, "pmOpc")
    printLog(f"pmOpc: {pmOpc}", console=False)

    global pmRead
    pmRead = readConfigFile(pm_driver, "pmRead")
    printLog(f"pmRead: {pmRead}", console=False)

    global pmConfigZS
    pmConfigZS = readConfigFile(pm_driver, "pmConfigZS")
    printLog(f"pmConfigZS: {pmConfigZS}", console=False)

    global pmTrigMeas
    pmTrigMeas = readConfigFile(pm_driver, "pmTrigMeas")
    printLog(f"pmTrigMeas: {pmTrigMeas}", console=False)

    global pmAvgQuery
    pmAvgQuery = readConfigFile(pm_driver, "pmAvgQuery")
    printLog(f"pmAvgQuery: {pmAvgQuery}", console=False)

    global pmAutoAvgOn
    pmAutoAvgOn = readConfigFile(pm_driver, "pmAutoAvgOn")
    printLog(f"pmAutoAvgOn: {pmAutoAvgOn}", console=False)

    # Get Step Attenuator Driver Commands
    global att_x_cmd_list
    global att_y_cmd_list
    att_x_cmd_list = readConfigFile(attenuator_driver, "x_chann", sFunc='listout')
    for index, item in enumerate(att_x_cmd_list):
        stripped = item.strip()
        att_x_cmd_list[index] = stripped
    printLog(f"att_x_cmd_list: {att_x_cmd_list}", console=False)

    att_y_cmd_list = readConfigFile(attenuator_driver, "y_chann", sFunc='listout')
    for index, item in enumerate(att_y_cmd_list):
        stripped = item.strip()
        att_y_cmd_list[index] = stripped
    printLog(f"att_y_cmd_list: {att_y_cmd_list}", console=False)

    return True


def initialize_visa_get_list():
    import pyvisa as visa
    global rm
    global list
    global resources
    global inst

    # searchString = searchString.lower()
    print('Refreshing VISA resource list...')
    rm = visa.ResourceManager()
    list = rm.list_resources()  # Place the resources into tuple "List"
    resources = []  # Prep for the tuple to be converted "resources" from "list"

    for i, a in enumerate(list):  # Enumerate through the tuple "list"
        resources.append(a)  # Place the current enumeration into the indexed spot of "resources"

    return resources


def set_visa_resource(device_name, search_resource_string='', perform_idn=True, idn_string='*IDN?'):

    printLog(f'Getting {device_name} resource string...')
    tempBool = False
    not_listed_string = '- Enter Resource String Manually'
    refresh_list = '- Refresh VISA Resource List'
    resource_string_no_idn = '- Enter Resource String And Skip IDN Confirmation'

    while tempBool == False:
        visa_device_list = initialize_visa_get_list()
        visa_device_list.append(not_listed_string)
        visa_device_list.append(resource_string_no_idn)
        visa_device_list.append(refresh_list)

        if search_resource_string == '':
            selected_resource = list_selection_box(visa_device_list, field1='Detected VISA Resource List',
                                                   field2=f'Choose the appropriate resource string for device: {device_name}',
                                                   window_title='Resource List',
                                                   width=60)
        else:
            selected_resource = search_resource_string
            formatted_search_rsrc_string = search_resource_string.lower()

            for item in visa_device_list:
                tempString = item.lower()

                if formatted_search_rsrc_string in tempString or tempString in formatted_search_rsrc_string:
                    selected_resource = item
                    search_resource_string = ''
                    break

        if selected_resource == not_listed_string or selected_resource == resource_string_no_idn:

            perform_idn = False if selected_resource == resource_string_no_idn else True

            tempString = f'Please type in the full resource string associated with device {device_name}\n'
            tempString += '(e.g., GPIB0::13::INSTR)'
            selected_resource = text_entry_box(title='Manual Visa Resource Entry', field1=tempString, field2='Resource:')

        if selected_resource == refresh_list:
            time.sleep(1)
        else:
            # print(selected_resource)
            try:
                inst = rm.open_resource(selected_resource)
                if perform_idn:
                    response = queryVisa(inst, idn_string, sFunc="", retryQty=5)
                    printLog(f'Set and queried {device_name} - Resource: {inst}, Responded: {response}')
                    tempString = f'Verify the selected resource for device >{device_name}< is correct:\n\n' \
                                 f'Resource Address: {inst}\n\n' \
                                 f'Received: {response}'
                    if debug_flag == False:
                        tempBool = yes_no_popup_simple(tempString)
                    else:
                        tempBool = True
                else:
                    tempString = f'Verify the selected resource for device >{device_name}< is correct, and the ' \
                                 f'device is connected and powered on \n\n' \
                                 f'Resource: {selected_resource}'
                    if debug_flag == False:
                        tempBool = yes_no_popup_simple(tempString)
                    else:
                        tempBool = True

                    if not tempBool:
                        selected_resource = ''
                        search_resource_string = ''
                    else:
                        printLog(f'User manually confirmed {device_name} - Resource: {inst}')

            except:
                printLog(f'Failed to get open resource string: {selected_resource}', errorInfo=True, console=False)
                tempString = f'The selected resource for remote instrument >{device_name}< is invalid!' \
                             f'\n\nResource: {selected_resource}\n\nPlease try again.'
                msg_box_simple(tempString)
                selected_resource = ''
                search_resource_string = ''

    return inst


def list_selection_box(input_list, field1='', field2='', window_title='', button_label='Select', width=150, height=30):

    import PySimpleGUI as sg

    """
        Allows you to "browse" through the Theme settings.  Click on one and you'll see a
        Popup window using the color scheme you chose.  It's a simple little program that also demonstrates
        how snappy a GUI can feel if you enable an element's events rather than waiting on a button click.
        In this program, as soon as a listbox entry is clicked, the read returns.
    """

    sg.theme('System Default')

    layout = [[sg.Text(field1)],
              [sg.Text(field2)],
              ### [sg.Listbox(values=sg.theme_list(), size=(150, 30), key='-LIST-', enable_events=True)],
              [sg.Listbox(values=input_list, size=(width, height), key='-LIST-', enable_events=True)],
              [sg.Button(button_label)]]

    window = sg.Window(window_title, layout, keep_on_top=True)

    while True:  # Event Loop
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Select'):
            # sg.theme(values['-LIST-'][0])
            # sg.popup_get_text('This is {}'.format(values['-LIST-'][0]))
            selected = format(values['-LIST-'][0])
            # print(f'selected: {selected}')
            break

    window.close()

    return selected


def text_entry_box(title='test', field1='field1', field2='field2'):
    import PySimpleGUI as sg

    sg.theme('System Default')  # Add some color to the window

    # Very basic window.  Return values using auto numbered keys

    layout = [
        [sg.Text(field1)],
        [sg.Text(field2, size=(15, 1)), sg.InputText()],
        # [sg.Text('Address', size=(15, 1)), sg.InputText()],
        # [sg.Text('Phone', size=(15, 1)), sg.InputText()],
        # [sg.Submit(), sg.Cancel()]
        [sg.Submit()]
    ]

    window = sg.Window(title, layout, keep_on_top=True)
    event, values = window.read()
    window.close()
    # print(event, values[0], values[1], values[2])  # the input data looks like a simple list when auto numbered
    return values[0]


def msg_box_simple(message_string):
    import PySimpleGUI as sg

    sg.theme('System Default')  # Add some color to the window

    sg.Popup(message_string, keep_on_top=True)


def yes_no_popup_simple(message_string):
    import PySimpleGUI as sg

    sg.theme('System Default')  # Add some color to the window

    window = sg.popup_yes_no(message_string, keep_on_top=True)
    tempBool = True if window == 'Yes' else False

    return tempBool


def yes_no_other_popup(message_string, other_str_text='  Skip ', btn_focus=1, window_title='Message Prompt', lineLength=40):
    import PySimpleGUI as sg
    y_focus = False
    n_focus = False
    o_focus = False

    btn_focus_type_check = f'{type(btn_focus)}'
    if not 'int' in btn_focus_type_check:
        btn_focus = 1

    if btn_focus == 1:
        y_focus = True
    elif btn_focus == 0:
        n_focus = True
    else:
        o_focus = True

    # lineLength = 40

    formatted_msg = ''
    cntr = 0
    for index, item in enumerate(message_string):
        formatted_msg += item
        cntr += 1

        if item == " ":
            spliced_str = message_string[index+1:]
            cntr2 = 0
            for item in spliced_str:
                if item == " ":
                    break
                cntr2+=1

            if cntr + cntr2 >= lineLength:
                formatted_msg += '\n'
                cntr = 0

        elif cntr >= lineLength:
            if not message_string[index+1] == " ":
                formatted_msg += '-\n'
            else:
                formatted_msg += '\n'
            cntr = 0


    sg.theme('System Default')

    layout = [[sg.Text(formatted_msg)],
              [sg.Text(' ' * lineLength)],
              [sg.Button('  Yes  ', key='-yes-', focus=y_focus), sg.Button('  No   ', key='-no-', focus=n_focus), sg.Button(other_str_text, key='-other-', focus=o_focus)],
              ]
    # Create the Window
    window = sg.Window(window_title, layout, keep_on_top=True, use_default_focus=False)

    while True:
        event, values = window.read()
        event_str = f'{event}'
        print(f'event: {event_str}')
        if 'yes' in event_str:
            return_val = 1
            break
        elif 'no' in event_str:
            return_val = 0
            break
        elif 'other' in event_str:
            return_val = -1
            break

        if event in (sg.WIN_CLOSED, 'Continue'):
            break
        if event == 'close':
            break

    window.close()

    return return_val


def file_browse_window(ext_description='ALL Files', ext_type='*.*', message='Click button below to choose file',
                       button_text='Browse'):
    import PySimpleGUI as sg

    sg.theme('System Default')

    file_browse_parameters = sg.FileBrowse(button_text = button_text,
        file_types = ((ext_description, ext_type),),
        initial_folder = None,
        )

    layout = [[sg.Text(message)],
              [sg.Input(key='-FILE-', visible=False, enable_events=True), file_browse_parameters]]

    event, values = sg.Window('File Compare', layout).read(close=True)

    return f'{values["-FILE-"]}'

def error_and_exit(messag='', pause_time=1):
    temp_string = f'Fatal Error!\n\n{messag}\n\nProgram will now terminate.'
    msg_box_simple(temp_string)
    printLog(temp_string)
    print(f'Closing in {pause_time} seconds...')
    time.sleep(pause_time)
    exit()


def return_class_type(variable):
    test = f'{type(variable)}'
    test = test.split()
    test = test[1]
    test = test[1:]
    test = test[:-2]
    return test


def sanitize_variable(input_variable, default_response='', specified_class='', error_response='error', eval_operation='',
                      eval_threshold=''):

    def evaluate_variable(variable_to_evaluate, operation_type, limit_value, corrected_value):
        return_value = variable_to_evaluate
        variable_class_type = return_class_type(variable_to_evaluate)

        if not variable_class_type in ['int', 'float']:
            return corrected_value

        if operation_type == 'at least':
            if variable_to_evaluate < limit_value:
                return_value = corrected_value
        elif operation_type == 'at most':
            if variable_to_evaluate > limit_value:
                return_value = corrected_value

        return return_value

    def cast_int(variable):
        return int(variable)

    def cast_str(variable):
        return str(variable)

    def cast_float(variable):
        return float(variable)

    reclass_variable = {'int' : cast_int,
                        'str' : cast_str,
                        'float' : cast_float}

    specified_class = specified_class.lower()
    eval_operation = eval_operation.lower()

    if specified_class == 'int' and return_class_type(error_response) != 'int':
        error_response = 999_999_999
    elif specified_class == 'float' and return_class_type(error_response) != 'float':
        error_response = 999_999_999.999

    new_value = input_variable

    class_type_list = ['str', 'int', 'float', 'list', 'tuple', '']
    eval_operation_list = ['at least', 'at most', '']

    if (not specified_class in class_type_list) or (not eval_operation in eval_operation_list):
        return error_response

    if specified_class != '':
        input_class_type = return_class_type(input_variable)

        if input_class_type != specified_class:
            try:
                new_value = reclass_variable[specified_class](input_variable)
            except:
                new_value = error_response

    if new_value == error_response and default_response != '':
        new_value = default_response
    elif new_value != error_response and eval_operation != '' and default_response != '' and eval_threshold != '':
        new_value = evaluate_variable(new_value, eval_operation, eval_threshold, default_response)

    return new_value


def import_txt_file(file_path):
    output_list = []

    if os.path.exists(file_path):
        with open(file_path, "r") as f:
            for item in f:
                cleaned_item = item.strip('\n')
                output_list.append(cleaned_item)

    return output_list


def get_attenuator_standard_data(linearityCalDataFilePath11, linearityCalDataFilePath110):

    def confirm_files_exist(att_file_path, att_name):
        # Verify the data file paths are valid
        ext = '*.csv'
        ext_desc = 'CSV Files'
        btn_msg = 'Browse Standard Data File'

        temp_msg = f'Specified filepath for the {att_name} standard data is unavailable or invalid.\n\n' \
                   'Please choose the appropriate file:'
        tempBool = file_check_exists(att_file_path)
        tempBool2 = False
        while tempBool == False :

            att_file_path = file_browse_window(ext_description=ext_desc, ext_type=ext,
                                               message=temp_msg,
                                               button_text=btn_msg)
            tempBool = file_check_exists(att_file_path)

        return att_file_path

    def confirm_serial_from_att_data(att_file, att_name):
        temp_msg = ''
        tempBool = False
        file_data_list = import_txt_file(att_file)

        model_serial_info = ''
        for item in file_data_list:
            item_lowered = item.lower()

            if 'model' in item_lowered:
                item_cleaned = item.replace(',', ': ')
                model_serial_info += f'{item_cleaned}\n'
            elif 'serial' in item_lowered:
                item_cleaned = item.replace(',', ': ')
                model_serial_info += item_cleaned

        if model_serial_info == '':
            temp_msg = f'Could not locate model or serial info for the {att_name} attenuator!\n\n' \
                       f'Please manually confirm the correct data file is selected:\n\n' \
                       f'{att_file}'
            # msg_box_simple(temp_msg)
        else:
            temp_msg = f'Please confirm the model and serial info obtained from the data file for the {att_name} ' \
                       f'is correct:\n\n{model_serial_info}'
            tempBool = yes_no_popup_simple(temp_msg)

            if tempBool == False:
                temp_msg = f'Incorrect {att_name} data file selected! User selected that the model and/or serial number obtained from the data file does not match the standard to be used.'
            else:
                temp_msg = ''

        return (tempBool, temp_msg)

    def confirm_files_not_same(file1, file2):
        error_msg = ''
        if file1 == file2:
            error_msg = f'the chosen attenuator data filse are the same! Please re-select the files and ensure ' \
                       f'each attenuator is assigned a unique data file'
            # msg_box_simple(error_msg)
            return (False, error_msg)
        else:
            return (True, error_msg)

    def extract_att_data(att_file):
        file_data_list = import_txt_file(att_file)

        model_serial_info = ''
        freq_start = False
        att_data_nested_list = []
        for item in file_data_list:
            item_lowered = item.lower()

            if freq_start:
                item_list = item.split(',')
                att_data_nested_list.append(item_list)

            if 'frequency' in item_lowered:
                freq_start = True

        return att_data_nested_list

    error_msg = ''
    att_data_list11 = []
    att_data_list110 = []

    if not debug_flag:
        not_same_sn = False
        sn_check1 = False
        sn_check2 = False
        att_name1 = '11 dB Attenuator'
        att_name2 = '110 dB Attenuator'

        # linearityCalDataFilePath11 = confirm_files_exist(linearityCalDataFilePath11, att_name1)
        sn_check1, error_msg = confirm_serial_from_att_data(linearityCalDataFilePath11, att_name1)
        if sn_check1 == False:
            return (att_data_list11, att_data_list110, error_msg)
        linearityCalDataFilePath11 = '' if sn_check1 == False else linearityCalDataFilePath11


        # linearityCalDataFilePath110 = confirm_files_exist(linearityCalDataFilePath110, att_name2)
        sn_check2, error_msg = confirm_serial_from_att_data(linearityCalDataFilePath110, att_name2)
        if sn_check2 == False:
            return (att_data_list11, att_data_list110, error_msg)
        linearityCalDataFilePath110 = '' if sn_check2 == False else linearityCalDataFilePath110

        not_same_sn, error_msg = confirm_files_not_same(linearityCalDataFilePath11, linearityCalDataFilePath110)
        if not_same_sn == False:
            return (att_data_list11, att_data_list110, error_msg)

    printLog(f'Importing Att. Std data from: {linearityCalDataFilePath11}')
    att_data_list11 = extract_att_data(linearityCalDataFilePath11)
    printLog(f'Importing Att. Std data from: {linearityCalDataFilePath110}')
    att_data_list110 = extract_att_data(linearityCalDataFilePath110)

    # Check to ensure the att data frequency matches the test frequency specified in the template file
    temp_list = att_data_list11[0]
    template_frequency = temp_list[0]

    if template_frequency != test_frequency:
        error_msg = f'The attenuator data frequency ({template_frequency}) does not match the test frequency specified in the DUT ' \
                   f'template file ({test_frequency})!'
        # error_and_exit(error_msg)

    return (att_data_list11, att_data_list110, error_msg)


def access_atten_value(att_data_list11=[], att_data_list110=[], desired_att_value=0, last_att_value=0):

    def access_att_sub_list_data(item_list, param=''):
        item_param_dictionary = {'freq': 0,
                            'nom': 1,
                            's21mag': 2,
                            's21unc' : 3,
                            's11mag' : 4,
                            's11phase' : 5,
                            '' : 2}
        value = item_list[item_param_dictionary[param]]

        try:
            value = int(value)
        except:
            try:
                value = float(value)
            except:
                value = 999_999_999

        return value

    def get_index_desired_value(att_data_11, att_data_110, desired_att_value):
        data_index1 = 0
        data_index2 = 0
        nominal = 0
        remainder = desired_att_value

        # Check to see if there is an exact match in the 11 dB step attenuator
        for index, item in enumerate(att_data_11):
            nominal = access_att_sub_list_data(item, param='nom')

            if desired_att_value == nominal:
                data_index2 = index
                nominal = access_att_sub_list_data(item, param='nom')
                remainder = desired_att_value - nominal
                break

        if remainder > 0:
            # Check to see if there is an exact match in the 110 dB step attenuator
            for index, item in enumerate(att_data_110):
                nominal = access_att_sub_list_data(item, param='nom')

                if desired_att_value == nominal:
                    data_index1 = index
                    nominal = access_att_sub_list_data(item, param='nom')
                    remainder = desired_att_value - nominal
                    break

        # If still no match is found, distribute between the two attenuators
        if remainder > 0:
            data_index1 = 0
            nominal = 0
            remainder = 0
            for index, item in enumerate(att_data_110):
                nominal = access_att_sub_list_data(item, param='nom')

                if nominal != 0:
                    diff = desired_att_value / nominal
                    # print(diff)
                    if diff <= 1:
                        temp_item = att_data_110[data_index1]
                        nominal = access_att_sub_list_data(temp_item, param='nom')
                        remainder = desired_att_value - nominal
                        break

                data_index1 = index

            if remainder == 0:
                temp_item = att_data_110[data_index1]
                nominal = access_att_sub_list_data(temp_item, param='nom')
                remainder = desired_att_value - nominal

            data_index2 = 0
            nominal = 0
            for index, item in enumerate(att_data_11):
                nominal2 = access_att_sub_list_data(item, param='nom')

                if nominal2 != 0:

                    if remainder / nominal2 < 1:
                        break

                data_index2 = index

        return (data_index1, data_index2)

    def generate_gpib_command(index_11, index_110, x_cmd_list, y_cmd_list):

        x_cmd = x_cmd_list[index_11]
        y_cmd = y_cmd_list[index_110]

        return f'{x_cmd}; {y_cmd}'

    def calc_att_val_and_unc(index_11, index_110, att_data_11, att_data_110):

        freq_11 = access_att_sub_list_data(att_data_11[index_11], param='freq')
        nom_11 = access_att_sub_list_data(att_data_11[index_11], param='nom')
        s21_11 = access_att_sub_list_data(att_data_11[index_11], param='s21mag')
        unc_11 = access_att_sub_list_data(att_data_11[index_11], param='s21unc')
        s11_11 = access_att_sub_list_data(att_data_11[index_11], param='s11mag')
        phase_11 = access_att_sub_list_data(att_data_11[index_11], param='s11phase')

        freq_110 = access_att_sub_list_data(att_data_110[index_110], param='freq')
        nom_110 = access_att_sub_list_data(att_data_110[index_110], param='nom')
        s21_110 = access_att_sub_list_data(att_data_110[index_110], param='s21mag')
        unc_110 = access_att_sub_list_data(att_data_110[index_110], param='s21unc')
        s11_110 = access_att_sub_list_data(att_data_110[index_110], param='s11mag')
        phase_110 = access_att_sub_list_data(att_data_110[index_110], param='s11phase')

        combined_att = s21_11 + s21_110
        combined_unc = math.sqrt((unc_11**2) + (unc_110**2))

        return_dict = {'freq_11' : freq_11,
                       'nom_11' : nom_11,
                       's21_11' : s21_11,
                       'unc_11' : unc_11,
                       's11_11' : s11_11,
                       'phase_11' : phase_11,
                       'freq_110' : freq_110,
                       'nom_110' : nom_110,
                       's21_110' : s21_110,
                       'unc_110' : unc_110,
                       's11_110' : s11_110,
                       'phase_110' : phase_110,
                       'combined_att' : combined_att,
                       'combined_unc' : combined_unc}
        return return_dict

    # parameter = parameter.lower()
    # allowed_parameters = ['freq', 'nom', 's21mag', 's21unc', 's11mag', 's11phase', '']
    # if not parameter in allowed_parameters:
    #     return 999_999_999

    if return_class_type(desired_att_value) != 'int':
        desired_att_value = sanitize_variable(desired_att_value, default_response='', specified_class='int',
                                              error_response='error',
                                                eval_operation='',
                                                eval_threshold='')
        if desired_att_value == 'error':
            return [999_999_999, 999_999_999, 999_999_999, 999_999_999, 999_999_999]

    if return_class_type(last_att_value) != 'int':
        last_att_value = sanitize_variable(last_att_value, default_response='', specified_class='int',
                                              error_response='error',
                                                eval_operation='',
                                                eval_threshold='')
        if last_att_value == 'error':
            return [999_999_999, 999_999_999, 999_999_999, 999_999_999, 999_999_999]

    if desired_att_value > 121:
        desired_att_value = 121
    elif desired_att_value < 0:
        desired_att_value = 0

    if last_att_value > 121:
        last_att_value = 121
    elif last_att_value < 0:
        last_att_value = 0

    # For troubleshooting the get_index_desired_value function
    # desired_att_value = 120
    index110_desired, index11_desired = get_index_desired_value(att_data_list11, att_data_list110, desired_att_value)
    # input(f'\n::: 11 Unit: {index11_desired}, 110 Unit: {index110_desired}')


    index110_last, index11_last = get_index_desired_value(att_data_list11, att_data_list110, last_att_value)

    # print(f'desired: {desired_att_value}, index110: {index110_desired}, index11: {index11_desired}')

    # Get the required GPIB command
    gpib_cmd = generate_gpib_command(index11_desired, index110_desired, att_x_cmd_list, att_y_cmd_list)
    # print(gpib_cmd)

    att_nom_data_desired = calc_att_val_and_unc(index11_desired, index110_desired, att_data_list11, att_data_list110)
    att_actual_desired = att_nom_data_desired['s21_11'] + att_nom_data_desired['s21_110']
    att_nom_data_last = calc_att_val_and_unc(index11_last, index110_last, att_data_list11, att_data_list110)
    att_actual_last = att_nom_data_last['s21_11'] + att_nom_data_last['s21_110']

    att_diff = att_nom_data_desired['combined_att'] - att_nom_data_last['combined_att']
    # combined_unc = math.sqrt((att_nom_data_desired['combined_unc']**2) + (att_nom_data_last['combined_unc']**2))
    combined_unc = att_nom_data_desired['combined_unc']
    return_list = [att_diff, combined_unc, gpib_cmd, att_actual_desired, att_actual_last]

    return return_list


def step_att_driver(resource_string, att_setting):

    if return_class_type(att_setting) != 'int':
        desired_att_level = sanitize_variable(att_setting, specified_class='int')
        if desired_att_level == 'error':
            temp_msg = '11713A Error set attenuator value class type error'
            printLog(temp_msg, console=True)
            return temp_msg

    att_param_list = access_atten_value(att_data_list11, att_data_list110, desired_att_value=att_setting)
    att_gpib_cmd = att_param_list[2]

    response = writeVisa(resource_string, att_gpib_cmd)

    # time.sleep(2)

    return response


def level_generator_and_power_meter_old(target_level, leveling_tol=0.02, settling_time=5, safe=True, max_output=20, retry_max=50):
    target_level = sanitize_variable(target_level, specified_class='float')
    max_output = sanitize_variable(max_output, specified_class='float')
    leveling_tol = sanitize_variable(leveling_tol, specified_class='float')
    settling_time = sanitize_variable(settling_time, specified_class='float')

    response = writeVisa(generator_resource, sGenOn)
    printLog(response)

    pm_level = queryVisa(pmeter_resource, pmRead)
    printLog(pm_level)
    pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)


    delta = target_level - pm_level
    abs_delta = abs(target_level - pm_level)
    sig_gen_current_level = queryVisa(generator_resource, sGenPowRead)
    sig_gen_current_level = sanitize_variable(sig_gen_current_level, specified_class='float', default_response=999_999_999)

    retry_counter = 0
    while abs_delta > leveling_tol:
        retry_counter+=1

        if retry_counter > retry_max:
            temp_msg = f'Sig Gen Leveling Routine Timeout!' \
                       f'\n\nLeveling loop exceeded the maximum loop limit ({retry_max})' \
                       f'\n\nPlease check connections, verify the generator can achieve the' \
                       f'\nleveled power goal of ({target_level:.4f}) dBm, and then continue ' \
                       f'\nto try again.'
            msg_box_simple(temp_msg)
            retry_counter = 0

        time.sleep(settling_time)

        pm_level = queryVisa(pmeter_resource, pmRead)
        pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)
        printLog(pm_level)

        delta = target_level - pm_level
        abs_delta = abs(target_level - pm_level)

        sig_gen_current_level = queryVisa(generator_resource, sGenPowRead)
        sig_gen_current_level = sanitize_variable(sig_gen_current_level, specified_class='float',
                                                  default_response=999_999_999)

        sig_gen_set_level = sig_gen_current_level + delta
        print(sig_gen_current_level)
        if safe:
            if not sig_gen_set_level > max_output:
                response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'{sig_gen_set_level}'))
            else:
                temp_msg = f'Sig Gen Set Level ({sig_gen_set_level}) Exceeded max safe output power ({max_output}) during the leveling routine!\n' \
                           f'Manually set the sensor to ({target_level}) dBm before continuing... '
                msg_box_simple(temp_msg)
        else:
            response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'{sig_gen_set_level}'))
        print(response)

        if debug_flag == True:
            print('-----------------------------------------')
            print('Generator / Power Meter Leveling Routine:')
            print(f'Target Level: {target_level}')
            print(f'Current Level: {pm_level}')
            print(f'Delta: {delta}')
            print(f'Abs Delta: {abs_delta}')
            print(f'Sig Gen Current Level: {sig_gen_current_level}')
            print(f'Sig Gen Set Level: {sig_gen_set_level}')
            print(f'Sig Gen Max Safe Level: {max_output}')

    return 'Generator / Power Meter Leveling Routine Complete!'


def level_generator_and_power_meter(target_level, leveling_tol=0.02, settling_time=5, safe=True, max_output=20, retry_max=50):

    def fine_level_generator_routine(target_level, sig_gen_set_level, loop_interval=0.75):
                                
        counter = 0
        leveling = True
        while leveling == True:
            time.sleep(loop_interval)
            response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'{sig_gen_set_level}'))
            
            sig_gen_current_level = queryVisa(generator_resource, sGenPowRead)
            sig_gen_current_level = sanitize_variable(sig_gen_current_level, specified_class='float',
                                          default_response=999_999_999)
            
            pm_level = queryVisa(pmeter_resource, pmRead)
            pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)
            # printLog(pm_level)
            
            delta_to_target = (target_level - pm_level)
            print(f"\nCounter: {counter}")
            print(f"Current Sig Gen Level: {sig_gen_current_level:.3f} dBm")
            print(f"Current PM Level: {pm_level:.3f} dBm")
            print(f"Target Level: {target_level:.3f} dBm")
            print(f"Delta to target: {delta_to_target:.3f} dB")
            
            if delta_to_target > 5:
                print("Incremented by 1 dB")
                sig_gen_set_level = sig_gen_current_level + 5
            elif delta_to_target > 1:
                print("Incremented by 1 dB")
                sig_gen_set_level = sig_gen_current_level + 1
            elif delta_to_target <= 1:
                print("Incremented by 0.1 dB")
                sig_gen_set_level = sig_gen_current_level + 0.1
                
            if delta_to_target < 0:
                print("Above Target Value")
                leveling = False
            elif abs(delta_to_target) <= 0.1:
                print("Incremented Counter")
                counter+=1
            
            if counter > 10:
                temp_msg = f'Sig Gen auto leveling routine has stalled!\n' \
                           f'Manually set the sensor to ({target_level}) dBm before continuing... '
                msg_box_simple(temp_msg)
                
            print(f"New Sig Gen Level: {sig_gen_set_level:.3f} dBm >...\n")
            
        return sig_gen_current_level


    target_level = sanitize_variable(target_level, specified_class='float')
    max_output = sanitize_variable(max_output, specified_class='float')
    leveling_tol = sanitize_variable(leveling_tol, specified_class='float')
    settling_time = sanitize_variable(settling_time, specified_class='float')

    response = writeVisa(generator_resource, sGenOn)
    printLog(response)

    pm_level = queryVisa(pmeter_resource, pmRead)
    printLog(pm_level)
    pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)


    delta = target_level - pm_level
    abs_delta = abs(target_level - pm_level)
    sig_gen_current_level = queryVisa(generator_resource, sGenPowRead)
    sig_gen_current_level = sanitize_variable(sig_gen_current_level, specified_class='float', default_response=999_999_999)

    retry_counter = 0
    while abs_delta > leveling_tol:
        retry_counter+=1

        if retry_counter > retry_max:
            temp_msg = f'Sig Gen Leveling Routine Timeout!' \
                       f'\n\nLeveling loop exceeded the maximum loop limit ({retry_max})' \
                       f'\n\nPlease check connections, verify the generator can achieve the' \
                       f'\nleveled power goal of ({target_level:.4f}) dBm, and then continue ' \
                       f'\nto try again.'
            msg_box_simple(temp_msg)
            retry_counter = 0

        time.sleep(settling_time)

        pm_level = queryVisa(pmeter_resource, pmRead)
        pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)
        printLog(pm_level)

        delta = target_level - pm_level
        abs_delta = abs(target_level - pm_level)

        sig_gen_current_level = queryVisa(generator_resource, sGenPowRead)
        sig_gen_current_level = sanitize_variable(sig_gen_current_level, specified_class='float',
                                                  default_response=999_999_999)

        sig_gen_set_level = sig_gen_current_level + delta
        print(f"Sig Gen Current Set Level{sig_gen_current_level:.3f} dBm, New Sig Get Set Level: {sig_gen_set_level:.3f} dBm, Max Safe Level: {max_output:.3f} dBm")
        if safe:
          
            if (sig_gen_set_level <= max_output):
                response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'{sig_gen_set_level}'))
            else:
                
                final_leveled_value = fine_level_generator_routine(target_level, sig_gen_current_level)
                max_output = final_leveled_value + 0.25
                    
                # temp_msg = f'Sig Gen Set Level ({sig_gen_set_level}) Exceeded max safe output power ({max_output}) during the leveling routine!\n' \
                #            f'Manually set the sensor to ({target_level}) dBm before continuing... '
                # msg_box_simple(temp_msg)
        else:
            response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'{sig_gen_set_level}'))
        print(response)

        if debug_flag == True:
            print('-----------------------------------------')
            print('Generator / Power Meter Leveling Routine:')
            print(f'Target Level: {target_level}')
            print(f'Current Level: {pm_level}')
            print(f'Delta: {delta}')
            print(f'Abs Delta: {abs_delta}')
            print(f'Sig Gen Current Level: {sig_gen_current_level}')
            print(f'Sig Gen Set Level: {sig_gen_set_level}')
            print(f'Sig Gen Max Safe Level: {max_output}')

    return 'Generator / Power Meter Leveling Routine Complete!'


def calc_uncertainty(nom_value, type_a_pct, drift_test_pct, sample_qty, attenuator_unc_pct=0, bias_sdev_pct=0):
    # Enter all uncertainty contributors as percent
    # convert the attenuator uncertainties to 1 Sigma

    # Convert percentages to nominal UOM
    attenuator_unc = nom_value / 100 * attenuator_unc_pct

    type_a = nom_value / 100 * type_a_pct

    drift_test = nom_value / 100 * drift_test_pct

    bias_sdev = nom_value / 100 * bias_sdev_pct


    # Get values to 1 Sigma as necessary

    attenuator_unc = attenuator_unc / 2

    # Combine all contributors at 1s
    sum_sq = (attenuator_unc ** 2) + (type_a ** 2) + (drift_test ** 2) + (bias_sdev ** 2)
    sum_contributors = attenuator_unc + type_a + drift_test + bias_sdev

    unc_1s = math.sqrt(sum_sq)

    # Lookup the coverage factor
    degrees_freedom = sample_qty - 1

    coverage_factor = Students_T_Lookup(degrees_freedom, Confidence=95.45)

    # Expand uncertainty to 2s
    unc_2s = unc_1s * coverage_factor

    # Perform uncertainty budget lookup
    unc_2s_pct = unc_2s / (nom_value / 100)
    nom_value_dBm = mW_dBm(nom_value)
    log.info(f'Uncertainty before budget lookup: {unc_2s_pct:.4f} %, {unc_2s:.2E} mW')
    new_unc_2s_pct = UncertaintyBudget.lookup(linBudgetTxtFile, unc_2s_pct, 50_000_000, nom_value_dBm)
    new_unc_2s = (nom_value / 100) * new_unc_2s_pct
    log.info(f'Uncertainty after budget lookup: {new_unc_2s_pct:.4f} %, {new_unc_2s:.2E} mW')

    temp_str = f'\n\nUnc_Eval: attenuator_unc: {attenuator_unc:.1E} mW ({(attenuator_unc / sum_contributors * 100):.2f}%), ' \
               f'degrees_freedom: {degrees_freedom:.0f}, ' \
               f'coverage_factor: {coverage_factor:.4f}, ' \
               f'type_a: {type_a:.2E} mW ({(type_a / sum_contributors * 100):.2f}%), ' \
               f'drift_test: {drift_test:.2E} mW ({(drift_test / sum_contributors * 100):.2f}%), ' \
               f'bias_sdev: {bias_sdev:.2E} mW ({(bias_sdev / sum_contributors * 100):.2f}%), ' \
               f'unc_1s: {unc_1s:.2E} mW, ' \
               f'unc_2s: {unc_2s:.2E} mW' \
               f'unc_2s (after lookup): {new_unc_2s:.2E} mW' \
               f'\nNote: percentages in parenthesis indicate what percent of the total uncertainty the item contributes.' \
               f'\n\nunc_2s: {unc_2s:.2E} mW, {(unc_2s / nom_value * 100):.2f}% of nominal' \
               f'\nunc_2s (after budget lookup): {new_unc_2s:.2E} mW, {(new_unc_2s / nom_value * 100):.2f}% of nominal\n'


    printLog(temp_str, console=verbose_flag)

    return new_unc_2s


def Pass_Fail_Eval(MsmtValue, LowerLimit, UpperLimit, uncertainty):
    Evaluation = ''
    FailFlag = False
    # Make sure the upper and lower limits aren't swapped
    temp_upper = UpperLimit
    temp_lower = LowerLimit
    if LowerLimit > UpperLimit:
        UpperLimit = temp_lower
        LowerLimit = temp_upper

    Pass_Upper = UpperLimit - uncertainty
    UGB1_Upper = UpperLimit
    UGB2_Upper = UpperLimit + uncertainty
    Pass_Lower = LowerLimit + uncertainty
    UGB1_Lower = LowerLimit
    UGB2_Lower = LowerLimit - uncertainty


    if MsmtValue <= Pass_Upper and MsmtValue >= Pass_Lower:
        Evaluation = "Pass"
    elif MsmtValue > UGB2_Upper or MsmtValue < UGB2_Lower:
        Evaluation = "Fail"
        FailFlag = True
    elif MsmtValue > Pass_Upper and MsmtValue <= UGB1_Upper:
        Evaluation = "UGB1"
    elif MsmtValue < Pass_Lower and MsmtValue >= UGB1_Lower:
        Evaluation = "UGB1"
    elif MsmtValue > UGB1_Upper and MsmtValue <= UGB2_Upper:
        Evaluation = "UGB2"
        FailFlag = True
    elif MsmtValue < UGB1_Lower and MsmtValue >= UGB2_Lower:
        Evaluation = "UGB2"
        FailFlag = True

    return (Evaluation, FailFlag)


def build_step_setting_dict(lin_step_list, ref_level):

    if debug_flag:
        print(f'build_step_setting_dict: ')

    writeListToFile(logFile, ["\nstep_setting_dict:\n"], write_type="a")

    index = 0
    step_setting_dict = {}
    for item in lin_step_list[::-1]:
        temp_msg = f'lin_step: {item}, att setting: {index}'
        writeListToFile(logFile, [temp_msg], write_type="a")
        if debug_flag:
            print(temp_msg)

        step_setting_dict[item] = index
        index += 1

    writeListToFile(logFile, ["\n"], write_type="a")
    if debug_flag:
        print('------------------\n\n')

    return step_setting_dict


def build_step_nominal_dict(step_setting_dict, ref_level):
    att_at_ref_setting = step_setting_dict[ref_level]
    # att_at_ref_setting = 10

    lin_step_list_backward = []
    for item in step_setting_dict:
        lin_step_list_backward.append(item)

    lin_step_list = []
    for item in reversed(lin_step_list_backward):
        lin_step_list.append(item)
    #   print(lin_step_list)

    writeListToFile(logFile, ["\nstep_nominal_dict & step_unc_dict:\n"], write_type="a")

    index = 0
    step_nominal_dict = {}
    step_unc_dict = {}
    for item in lin_step_list:
        att_setting = step_setting_dict[item]
        att_data_list = access_atten_value(att_data_list11=att_data_list11, att_data_list110=att_data_list110,
                                  desired_att_value=att_setting, last_att_value=att_at_ref_setting)
        # print(att_data_list)
        att_actual = ref_level - att_data_list[0]
        att_unc = att_data_list[1]
        temp_msg = f'lin_step: {item}, att_rel_ref: {att_actual}, unc: {att_unc}, att_set: {att_setting}, att_ref_set: {att_at_ref_setting}'
        if debug_flag:
            print(temp_msg)
        writeListToFile(logFile, [temp_msg], write_type="a")
        step_nominal_dict[item] = att_actual
        step_unc_dict[item] = att_unc
        index += 1
    print("")

    writeListToFile(logFile, ["\n"], write_type="a")


    return (step_nominal_dict, step_unc_dict)


def create_pscalcorr_lin_dat_file(csv_data_path, ps_cal_path):
    import pathlib
    import os

    global verbose_flag



    file_extension = pathlib.Path(csv_data_path).suffix
    filename_wout_extension = csv_data_path.strip(file_extension)
    new_filename = f'{filename_wout_extension}.dat'
    new_filename = standardize_file_path_format(new_filename)
    new_filename = rename_if_file_exists(new_filename)

    head, tail = os.path.split(new_filename)
    ps_cal_full_path = f'{ps_cal_path}/{tail}'
    ps_cal_full_path = standardize_file_path_format(ps_cal_full_path)
    ps_cal_full_path = rename_if_file_exists(ps_cal_full_path)

    textfile_list = import_txt_file(csv_data_path)

    # Dat file format:
    # Nom,   Measured (pct rel nominal), Limit, Unc, P/F
    #  -1, -1.001310826941515 (-0.030%),   0.5,   0, Pass

    # Delete the header line from the data list
    del textfile_list[0]
    dat_file_list = []
    dat_file_list.append(f"Procedure {cs_number}, Rev. {program_version}")

    for item in textfile_list:
        item_list = item.split(",")

        dBm_nom = float(item_list[1])
        mW_nom = dBm_mW(dBm_nom)
        dut_msd_dBm = float(item_list[2])
        dut_msd_mW = float(item_list[4])
        ul_mW = float(item_list[5])
        unc_mW = float(item_list[6])
        PF_Bool = item_list[8]
        # tol_pct = abs(((mW_nom / ul_mW) * 100) - 100)
        tol_pct = item_list[11]

        pct_of_nom = abs((dut_msd_mW / mW_nom * 100) - 100)

        if dut_msd_mW >= mW_nom:
            sign = '+'
        else:
            sign = '-'

        unc_pct = unc_mW / (mW_nom / 100)
        unc_pct = setSigDigits(unc_pct, 2)

        pf_flag = item_list[8]
        pf_flag = pf_flag.lower()
        if pf_flag == 'false':
            pf_flag = "Pass"
        else:
            pf_flag = "Fail"

        # print(f'msd: {dut_msd_mW} nominal: {mW_nom}')
        new_line = f'{dBm_nom:.3f},{dut_msd_dBm} ({sign}{pct_of_nom:.3f}%),{tol_pct},{unc_pct},{pf_flag}'
        if verbose_flag:
            print(new_line)
        dat_file_list.append(new_line)

    writeListToFile(new_filename, dat_file_list, write_type='a')

    return (new_filename, ps_cal_full_path)


class UncertaintyBudget:
    @staticmethod
    def GetUOM(budgetTxtFile):
        f = open(budgetTxtFile, 'r')
        budgetData = f.readlines()
        f.close()

        for line in budgetData:
            searchLine = line.lower()
            searchLine = searchLine.strip()
            if "uom" in searchLine:
                searchLine = searchLine.split(",")
                return searchLine[1]
        return "Unknown"

    @staticmethod
    def check_is_expired(budgetTxtFile):
        import datetime
        from datetime import date
        # Place contents of XML files into variable
        f = open(budgetTxtFile, 'r')
        budgetData = f.readlines()
        f.close()

        # Strip any existing whitespace out of the list elements
        for index, element in enumerate(budgetData):
            if element == "\n":
                del budgetData[index]
            else:
                newElement = element.strip()
                budgetData[index] = newElement

        # Placed current date into variable
        today = date.today()
        currentDate = today.strftime("%Y-%m-%d")

        # Obtain the line containing the expiration date of the budget file
        tempList = budgetData[0]
        tempList = tempList.strip()
        tempList = tempList.split(",")
        fileDate = tempList[1]
        fileDate = fileDate.split("-")  # Store date elements into a list
        currentDate = currentDate.split("-")  # Also store today's date elements into list

        # Convert dates into formats which can be compared against
        fileDateInFormat = datetime.datetime(int(fileDate[0]), int(fileDate[1]), int(fileDate[2]))
        currentDateInFormat = datetime.datetime(int(currentDate[0]), int(currentDate[1]), int(currentDate[2]))

        # Check to see if the file date has been exceeded
        if currentDateInFormat > fileDateInFormat:
            return True
        else:
            return False

    @staticmethod
    def lookup(budgetTxtFile, uncVal, uncFreq, uncMsmt=0.0):
        import datetime
        from datetime import date

        # Suppress scientific notation (this code ended up being unnecessary)
        # uncVal = f'{uncVal:.20f}'                       # 20 digits, because why not?
        # uncFreq = f'{uncFreq:.0f}'                      # Frequency in Hz should have no resolution after the decimal

        # Place contents of XML files into variable
        f = open(budgetTxtFile, 'r')
        budgetData = f.readlines()
        f.close()

        # Strip any existing whitespace out of the list elements
        for index, element in enumerate(budgetData):
            if element == "\n":
                del budgetData[index]
            else:
                newElement = element.strip()
                budgetData[index] = newElement

        # Placed current date into variable
        today = date.today()
        currentDate = today.strftime("%Y-%m-%d")

        # Obtain the line containing the expiration date of the budget file
        tempList = budgetData[0]
        tempList = tempList.strip()
        tempList = tempList.split(",")
        fileDate = tempList[1]
        fileDate = fileDate.split("-")  # Store date elements into a list
        currentDate = currentDate.split("-")  # Also store today's date elements into list

        # Convert dates into formats which can be compared against
        fileDateInFormat = datetime.datetime(int(fileDate[0]), int(fileDate[1]), int(fileDate[2]))
        currentDateInFormat = datetime.datetime(int(currentDate[0]), int(currentDate[1]), int(currentDate[2]))

        # Check to see if the file date has been exceeded
        if currentDateInFormat > fileDateInFormat:
            return "File > {} < expiration date is exceeded!".format(budgetTxtFile)

        # del budgetData[0]  # Delete the list element containing the date
        # del budgetData[0]  # Delete the list element containing the UOM

        # Get rid of all lines from the budget data which are not for performing the lookup
        for index, line_data in enumerate(budgetData):
            if not ">" in line_data:
                del budgetData[index]

        # Check to see if this budget format includes Power ranges
        tempList = budgetData[0].split(",")
        tempQty = len(tempList)
        if tempQty > 2:
            pRangePresent = True
        else:
            pRangePresent = False

        # Break out the elements of the uncertainty budget list
        freqList = []
        rangeList = []
        uncList = []
        for index, element in enumerate(budgetData):
            tempList = element.split(",")
            if pRangePresent == True:
                freqList.append(tempList[0])
                rangeList.append(tempList[1])
                uncList.append(float(tempList[2]))
            else:
                freqList.append(tempList[0])
                uncList.append(float(tempList[1]))

        # print(freqList)
        # print(rangeList)
        # print(uncList)

        # Find the current uncertainty frequency within the budget frequency list
        for index, element in enumerate(freqList):
            tempList = element.split(">")
            startFreq = float(tempList[0])
            stopFreq = float(tempList[1])
            uncFreq = float(uncFreq)

            if (uncFreq >= startFreq) and (uncFreq <= stopFreq):

                if pRangePresent == True:
                    tempList = rangeList[index].split(">")
                    startPow = float(tempList[0])
                    stopPow = float(tempList[1])

                    if (uncMsmt >= startPow) and (uncMsmt <= stopPow):
                        tempUncListVal = uncList[index]

                        if uncVal < tempUncListVal:
                            uncVal = tempUncListVal
                            return uncVal
                else:
                    tempUncListVal = uncList[index]
                    if uncVal < tempUncListVal:
                        uncVal = tempUncListVal
                        return uncVal

        return uncVal


def plot_data(input_filepath, output_filepath, plt_x=15, plt_y=8):
    import matplotlib.pyplot as plt
    import math

    input_data_list = readTxtFile(input_filepath)

    for index, item in enumerate(input_data_list):
        item = item.strip()
        input_data_list[index] = item

    # Delete first line of the data, which contains the header info
    del input_data_list[0]

    nominal_step = []
    nominal_val = []
    msmt_val = []
    upper_limit = []
    lower_limit = []
    upper_unc = []
    lower_unc = []
    for item in input_data_list:
        temp_list = item.split(",")
        # print(temp_list)

        step = float(temp_list[0])

        mw_lower = float(temp_list[3])
        mw_msd = float(temp_list[4])
        mw_upper = float(temp_list[5])
        mw_unc = float(temp_list[6])
        mw_nom = (mw_lower + mw_upper) / 2

        dbm_lower = 10*math.log10(mw_lower)
        dbm_msd = 10 * math.log10(mw_msd)
        dbm_upper = 10 * math.log10(mw_upper)
        dbm_nom = 10 * math.log10(mw_nom)

        dbm_unc_lower = 10 * math.log10((mw_msd - mw_unc))
        dbm_unc_upper = 10 * math.log10((mw_msd + mw_unc))

        msd_normalized = dbm_msd - dbm_nom

        limit_normalized_lower = dbm_lower - dbm_nom
        limit_normalized_upper = dbm_upper - dbm_nom

        unc_normalized_lower = dbm_unc_lower - dbm_nom
        unc_normalized_upper = dbm_unc_upper - dbm_nom

        nominal_step.append(step)
        nominal_val.append(0)
        msmt_val.append(msd_normalized)
        lower_limit.append(limit_normalized_lower)
        upper_limit.append(limit_normalized_upper)
        lower_unc.append(unc_normalized_lower)
        upper_unc.append(unc_normalized_upper)


    plot_filename = output_filepath
    plt.rcParams["figure.figsize"] = (plt_x, plt_y)
    ax = plt.gca()
    ax.grid(True)
    plt.plot(nominal_step, nominal_val, '.', label="Nominal", color='green')
    plt.plot(nominal_step, msmt_val, label="DUT Msd.")
    plt.plot(nominal_step, upper_limit, 'b-', label="Limit", color='red')
    plt.plot(nominal_step, lower_limit, 'b-', color='red')
    plt.plot(nominal_step, upper_unc, '-.', label="Unc", color='orange')
    plt.plot(nominal_step, lower_unc, '-.', color='orange')
    plt.xticks(rotation=-90)
    # naming the x axis
    plt.xlabel('Linearity Step (dBm)')
    plt.xticks(nominal_step)
    # naming the y axis
    plt.ylabel('Delta Nominal (dB)')
    # giving a title to my graph
    plt.title(f'Linearity Test Result\n'
              f'File: {output_filepath}')

    # show a legend on the plot
    plt.legend()

    # Save the smoothing plot
    plt.savefig(plot_filename)
    # plt.show()
    try:
        plt.cla()
    except:
        pass
    try:
        plt.clf()
    except:
        pass
    try:
        plt.close()
    except:
        pass

    return plot_filename


def calc_resol_qty(input_float, additonal_res=4, debug=False):
    try:
        input_float = float(input_float)
    except:
        print(f"Input value {input_float} is not a float!")
        return 0

    input_str = f'{input_float}'

    cntr = 0
    found_period = False
    for character in input_str:
        if debug:
            print(character)
        if not (character == "0" or character == ".") and found_period == False:
            break

        elif character == ".":
            found_period = True

        elif found_period:
            if character == "0":
                cntr+=1
            else:
                break

    if debug:
        print(f'Found {cntr} zeroes')
    res_qty = cntr + additonal_res
    return (res_qty)


class GuiProgramWindow:
    # Requires the PySimpleGUI module to be installed
    import PySimpleGUI as sg
    import threading as thrd
    import time

    def __init__(self, windowTitle, ConsoleStringList=[], consoleQuantityOfLines=20, inputString=''):
        self.windowTitle = windowTitle
        self.ConsoleStringList = ConsoleStringList
        self.consoleQuantityOfLines = consoleQuantityOfLines
        self.inputString = inputString
        # self.DUT_File_Type = "All Files (*.*)"

        # Automatically Declared Variables
        self.DutFileExtension = "*.lin"
        self.DutFileDescription = "Cal Template File"
        self.StdFileExtension = "*.csv"
        self.StdFileDescription = "N7800 CSV File"
        self.dut_template_filepath = ''
        self.att_11_data_filepath = ''
        self.att_110_data_filepath = ''
        self.att_visa_string = ''
        self.pm_visa_string = ''
        self.gen_visa_string = ''
        self.DutDataFolder = None
        self.StdDataFolder = None
        self.StartLinearityTest = False
        self.CloseBool = False
        self.step_att_name = ''
        self.generator_name = ''
        self.pMeterName = ''
        self.stepAttVisaResourceIdent = ''
        self.sGenVisaResource = ''
        self.pMeterVisaResourceIdent = ''
        self.set_attenuator_resource = None
        self.set_generator_resource = None
        self.set_pmeter_resource = None
        self.att_data_initial_path11 = ''
        self.att_data_initial_path110 = ''
        self.att_data_list11 = None
        self.att_data_list110 = None
        self.dut_asset_number = ''
        self.thread1 = None
        self.retest = None
        self.drift_test = None

        # self.LinFileDescription = "Test Desc."
        # self.LinFileExtension = "*.Dat"
        # self.ZsFileDescription = "Test Desc."
        # self.ZsFileExtension = "*.ZSC"
        # self.DutDataFolder = None
        # self.StdDataFolder = None
        # self.LinearityDataFileFolder = None
        # self.ZeroSetDataFileFolder = None
        # self.SelectedStandardsList = []
        # self.StartCorrectionProcess = False
        # self.DutFilePath = None
        # self.StandardFilePath = None
        # self.LinFilePath = None
        # self.ZeroSetFilePath = None
        # self.LinearityDataAvailable = False
        # self.ZeroSetDataAvailable = False
        # self.InputStandardsList = []
        # self.BackupStandardsList = []
        # self.thread1 = None
        # self.CloseBool = False

        self.sg.theme('SystemDefaultForReal')

    def _thread_function(self):
        # import PySimpleGUI as sg
        # import functools
        inputList = []

        # Set the File Browser File Type
        self.DUT_File_Type = [[self.DutFileDescription, self.DutFileExtension], ]
        self.STD_File_Type = [[self.StdFileDescription, self.StdFileExtension], ]
        # self.LIN_File_Type = [[self.LinFileDescription, self.LinFileExtension], ]
        # self.ZS_File_Type = [[self.ZsFileDescription, self.ZsFileExtension], ]

        outputList = []
        stringOutputList = []
        stringInputList = []
        for index, i in enumerate(inputList):
            stringInputList.append("{:15}".format(i))

        # Construct the GUI interface
        # self.sg.theme('SystemDefaultForReal')  # Sets the system default theme

        lineLength = 120
        bt = {'size': (7, 2)}
        lb = {'size': (100, 20), 'enable_events': (True), 'font': ('Courier 10')}
        fb = {'size': (115, 20)}
        resource_field = {'size': (80, 20)}
        resource_button = {'size': (23, 1)}

        header = {'font': ('Courier 10')}

        menu_def = [
            ['Misc', 'About...'], ]

        layout = [[self.sg.Menu(menu_def, tearoff=False)],
                  [self.sg.Text('_' * lineLength)],
                  [self.sg.Text('DUT Asset Number:'),
                   self.sg.In(key='-dut_asset_in-', **{'size': (7, 20)}, disabled=False, focus=True)],
                  [self.sg.FileBrowse('Browse DUT Cal Template File', target='-dutFile-', file_types=self.DUT_File_Type,
                                      initial_folder=self.DutDataFolder)],
                  [self.sg.In(key='-dutFile-', enable_events=True, **fb)],
                  # [self.sg.Button('Load DUT Cal Template',  key='-load_template_btn-')],
                  [self.sg.Text('_' * lineLength)],
                  [self.sg.Text('Set Cal Standards Data:')],
                  [self.sg.FileBrowse('Browse 11 dB Attenuator Data File', target='-stdFile11-',
                                      file_types=self.STD_File_Type, initial_folder=self.StdDataFolder, disabled=True)],
                  [self.sg.In(self.att_data_initial_path11, key='-stdFile11-', **fb, disabled=True)],
                  [self.sg.FileBrowse('Browse 110 dB Attenuator Data File', target='-stdFile110-',
                                      file_types=self.STD_File_Type, initial_folder=self.StdDataFolder, disabled=True)],
                  [self.sg.In(self.att_data_initial_path110, key='-stdFile110-', **fb, disabled=True)],
                  [self.sg.Text('_' * lineLength)],
                  [self.sg.Text('Configure Remote Resources:')],
                  [self.sg.Button('Confirm Attenuator Resource', **resource_button, key='-att_rsrc_btn-',
                                  disabled=True), self.sg.In(key='-att_rsrc-', **resource_field, disabled=True),
                   self.sg.Text('        ', key='-att_check-')],
                  [self.sg.Button('Confirm Generator Resource', **resource_button, key='-gen_rsrc_btn-', disabled=True),
                   self.sg.In(key='-gen_rsrc-', **resource_field, disabled=True),
                   self.sg.Text('        ', key='-gen_check-')],
                  [self.sg.Button('Confirm Power Meter Resource', **resource_button, key='-pm_rsrc_btn-',
                                  disabled=True), self.sg.In(key='-pm_rsrc-', **resource_field, disabled=True),
                   self.sg.Text('        ', key='-pm_check-')],
                  [self.sg.Text('_' * lineLength)],
                  [self.sg.Button('Start Linearity Calibration', font=('Helvetica', 10, 'bold'), disabled=True,
                                  key='-StartCal-'),
                   self.sg.Checkbox('Auto-Retest UGB & Failed Points', default=True, key="-retest-"),
                   self.sg.Checkbox('Include 30 Minute System Drift', default=False, key="-drift_test-"),
                   self.sg.Text(
                       '                                                                                      ',
                       key="-TF2-")],
                  # [self.sg.Button('Select Standards'), self.sg.Text('(Use this button to choose calibration standards)', key="-TF1-")],
                  # [self.sg.Button('Start Correction Process', font=('Helvetica', 10, 'bold'), disabled=False, key='-StartCorr-'), self.sg.Text('                                                                                      ', key="-TF2-")],
                  # [self.sg.Text('_' * lineLength)],
                  # [self.sg.Text('Optional Calibration Data:')],
                  # [self.sg.FileBrowse('Browse Linearity Data', target='-linFile-', file_types=self.LIN_File_Type, initial_folder=self.LinearityDataFileFolder, disabled=False, key='-LinBrowseBtn-',)],
                  # [self.sg.In(key='-linFile-', **fb)],
                  # [self.sg.FileBrowse('Browse Zero Set Data', target='-zeroSetFile-', file_types=self.ZS_File_Type, initial_folder=self.ZeroSetDataFileFolder,
                  #                     disabled=False, key='-SeroSetBrowseBtn-')],
                  # [self.sg.In(key='-zeroSetFile-', **fb)],
                  # [self.sg.Text('_' * lineLength)],
                  # [self.sg.Text('Console Output')],
                  # [self.sg.Output(size=(115, 15), key="-console-")]
                  # # [self.sg.Listbox(values=inputList, key='-console-', **lb)]
                  ]
        # Create the Window
        self.window = self.sg.Window(self.windowTitle, layout, keep_on_top=True)

        while True:
            event, values = self.window.read()
            if self.CloseBool == False:
                self.dut_asset_number = values['-dut_asset_in-']
                self.dut_template_filepath = values['-dutFile-']
                self.att_11_data_filepath = values['-stdFile11-']
                self.att_110_data_filepath = values['-stdFile110-']
                self.att_visa_string = values['-att_rsrc-']
                self.gen_visa_string = values['-gen_rsrc-']
                self.pm_visa_string = values['-pm_rsrc-']
                self.retest = values['-retest-']
                self.drift_test = values['-drift_test-']

            event_var_type = return_class_type(event)
            # print(f"Class Type: {return_class_type(event)}")
            if not "NoneType" in event_var_type:
                if event in ('-dutFile-'):
                    self.load_dut_template()

            if event == 'About...':
                self.about_menu_selection()
            if self.CloseBool:
                print("Closed Bool True")
                break
            if event in (self.sg.WIN_CLOSED, 'Continue'):
                break
            # if event in ('-load_template_btn-'):
            #     self.load_dut_template()
            if event in ('-att_rsrc_btn-'):
                self.create_att_visa()
            if event in ('-gen_rsrc_btn-'):
                self.create_gen_visa()
            if event in ('-pm_rsrc_btn-'):
                self.create_pm_visa()
            if event == 'close':
                break
            if event in ('-StartCal-'):
                print('Start Cal Button')
                self.Start_Lin_Test()

        self.window.close()

    def load_dut_template(self):
        print('Load DUT template')

        if len(self.dut_template_filepath) == 0:
            msg_box_simple("You must browse and select the DUT Cal Template File First!")
        elif not file_check_exists(self.dut_template_filepath):
            msg_box_simple(f"The template file does not exist!\n\nTried: {self.dut_template_filepath}")
            self.window['-dutFile-'].update('')
        else:
            successful_template_load = get_dut_template_data(self.dut_template_filepath)

            if not successful_template_load:
                self.dut_template_filepath = ''
                self.window['-dutFile-'].update('')
                return False

            self.step_att_name = step_att_name
            self.generator_name = generator_name
            self.pMeterName = pMeterName
            self.stepAttVisaResourceIdent = stepAttVisaResourceIdent
            self.sGenVisaResource = sGenVisaResource
            self.pMeterVisaResourceIdent = pMeterVisaResourceIdent

            # Unlock all the fields
            self.lock_unlock_lower_fields(locked=False)

            print(f'Current DUT template file: >{self.dut_template_filepath}<')

    def create_att_visa(self):
        print('Create Att VISA')
        self.set_attenuator_resource = set_visa_resource(self.step_att_name,
                                                         search_resource_string=self.stepAttVisaResourceIdent,
                                                         perform_idn=False)
        self.window['-att_rsrc-'].update(self.set_attenuator_resource)

    def create_gen_visa(self):
        print('Create Generator VISA')
        self.set_generator_resource = set_visa_resource(self.generator_name,
                                                        search_resource_string=self.sGenVisaResource)
        self.window['-gen_rsrc-'].update(self.set_generator_resource)

    def create_pm_visa(self):
        print('Create PM VISA')
        self.set_pmeter_resource = set_visa_resource(self.pMeterName,
                                                     search_resource_string=self.pMeterVisaResourceIdent)
        self.window['-pm_rsrc-'].update(self.set_pmeter_resource)

    def Start_Lin_Test(self):

        loaded_att_successfully = False
        try:

            self.att_data_list11, self.att_data_list110, error_msg = get_attenuator_standard_data(
                self.att_11_data_filepath,
                self.att_110_data_filepath)

            if error_msg == '':
                loaded_att_successfully = True

        except Exception as error:
            error_msg = f'{error}'

        if len(self.dut_asset_number) == 0:
            msg_box_simple(f"You must enter the DUT asset number before proceeding!")
        elif len(self.att_11_data_filepath) == 0:
            msg_box_simple(f"Select the 11 dB attenuator N7800 data file before proceeding!")
            self.window['-stdFile11-'].update('')
        elif not file_check_exists(self.att_11_data_filepath):
            msg_box_simple(f"The 11 dB attenuator file does not exist!\n\nTried: {self.att_11_data_filepath}")
            self.window['-stdFile11-'].update('')
        elif len(self.att_110_data_filepath) == 0:
            msg_box_simple(f"Select the 110 dB attenuator N7800 data file before proceeding!")
            self.window['-stdFile110-'].update('')
        elif not file_check_exists(self.att_110_data_filepath):
            msg_box_simple(f"The 11 dB attenuator file does not exist!\n\nTried: {self.att_110_data_filepath}")
            self.window['-stdFile110-'].update('')
        elif self.att_110_data_filepath == self.att_11_data_filepath:
            msg_box_simple(f"The 11 dB attenuator data file cannot be the same as the 110 dB attenuator data file!")
            self.window['-stdFile11-'].update('')
            self.window['-stdFile110-'].update('')
        elif loaded_att_successfully == False:
            msg_box_simple(
                f"Failed to load the 11 and 110 dB attenuator data files. Error:\n\n{error_msg}\n\nPlease try again!")
            self.window['-stdFile11-'].update('')
            self.window['-stdFile110-'].update('')
        elif len(self.att_visa_string) == 0:
            msg_box_simple(f"Select the attenuator driver remote resource string before proceeding!")
        elif len(self.gen_visa_string) == 0:
            msg_box_simple(f"Select the signal generator remote resource string before proceeding!")
        elif len(self.pm_visa_string) == 0:
            msg_box_simple(f"Select the power meter remote resource string before proceeding!")
        else:
            self.lock_unlock_upper_fields(locked=True)
            self.lock_unlock_lower_fields(locked=True)
            self.window['-TF2-'].update('(Linearity Cal In Process...)', font=('Helvetica', 10, 'bold'))
            self.StartLinearityTest = True
            self.close_window()

    def lock_unlock_upper_fields(self, locked=False):
        self.window.FindElement('-dutFile-').Update(disabled=locked)
        self.window.FindElement('Browse DUT Cal Template File').Update(disabled=locked)
        # self.window.FindElement('-load_template_btn-').Update(disabled=locked)

    def lock_unlock_lower_fields(self, locked=False):
        self.window.FindElement('Browse 11 dB Attenuator Data File').Update(disabled=locked)
        self.window.FindElement('-stdFile11-').Update(disabled=locked)
        self.window.FindElement('Browse 110 dB Attenuator Data File').Update(disabled=locked)
        self.window.FindElement('-stdFile110-').Update(disabled=locked)
        self.window.FindElement('-att_rsrc_btn-').Update(disabled=locked)
        self.window.FindElement('-att_rsrc-').Update(disabled=locked)
        self.window.FindElement('-gen_rsrc_btn-').Update(disabled=locked)
        self.window.FindElement('-gen_rsrc-').Update(disabled=locked)
        self.window.FindElement('-pm_rsrc_btn-').Update(disabled=locked)
        self.window.FindElement('-pm_rsrc-').Update(disabled=locked)
        self.window.FindElement('-StartCal-').Update(disabled=locked)
        # self.window['-TF2-'].update('(Linearity Cal In Process...)', font=('Helvetica', 10, 'bold'))

    def about_menu_selection(self):
        temp_str = f'You must be pretty bored to press the \"About\" button. Sorry things aren\'t more exciting right now.' \
                   f'\n\n' \
                   f'So I said in my heart, \"As death happens to the fool, death also happens to me, so why was ' \
                   f'I then more wise than a fool?\" Then I said in my heart, \"This also is vanity.\" For there is ' \
                   f'no more remembrance of the wise than of the fool forever, since all that now is will be forgotten' \
                   f' in the days to come. And how does a wise man die? The same way as the fool! Therefore I hated ' \
                   f'life because the work that was done under the sun was distressing to me, for all is vanity and ' \
                   f'grasping for the wind. Then I hated all my labor in which I had toiled under the sun, because I ' \
                   f'must leave it to the man who will come after me. And who knows whether he will be wise or a fool? ' \
                   f'Yet he will rule over all my labor in which I toiled and in which I have shown myself wise under ' \
                   f'the sun. This also is vanity. Therefore I turned my heart and despaired of all the labor in which' \
                   f' I had toiled under the sun. For there is a man whose labor is with wisdom, knowledge, and skill;' \
                   f' yet he must leave his heritage to a man who has not labored for it. This also is vanity and' \
                   f' a great evil. For what has man for all his labor, and for the striving of his heart with which' \
                   f' he has toiled under the sun? For all his days are sorrowful, and his work burdensome; even in' \
                   f' the night his heart takes no rest. This also is vanity... For what happens to the sons of men ' \
                   f'also happens to animals; one thing befalls them: as one dies, so dies the other. Surely, they all' \
                   f' have one breath; man has no advantage over animals, for all is vanity. All go to one place: all' \
                   f' are from the dust, and all return to dust. Who knows the spirit of the sons of men, which goes' \
                   f' upward, and the spirit of the animal, which goes down to the earth? So I perceived that nothing' \
                   f' is better than that a man should rejoice in his own works, for that is his heritage. For who can ' \
                   f'bring him to see what will happen after him?... For who knows what is good for man in life, all' \
                   f' the days of his vain life in which he passes like a shadow? Who can tell a man what will happen' \
                   f' after him under the sun?... Let us hear the conclusion of the whole matter: Fear God and keep His' \
                   f' commandments, for this is man’s all. For God will bring every work into judgment, including every' \
                   f' secret thing, whether good or evil.\n\n- Selections from Ecclesiastes'

        yes_no_other_popup(temp_str, other_str_text='  N/A ', btn_focus=1, window_title='The About Window...',
                           lineLength=120)

    def yes_no_popup_simple(self, message_string):

        self.sg.theme('System Default')  # Add some color to the window

        layout = [[self.sg.Button('Yes', key='-yes-'), self.sg.Button('No', key='-no-')]]

        self.window2 = self.sg.Window(self.windowTitle, layout, keep_on_top=True)

        while True:
            event, values = self.window2.read()

            if self.CloseBool:
                print("Closed Bool True")
                break
            if event in (self.sg.WIN_CLOSED, 'Continue'):
                break
            if event in ('-yes-'):
                print("Yes Button")

        window = self.sg.popup_yes_no(message_string, keep_on_top=True)
        tempBool = True if window == 'Yes' else False

        return tempBool

    # ======================================================================================================
    # def CheckStartCorrection(self):
    #     tempBool = True
    #     if self.DutFilePath == "":
    #         self.sg.popup('No DUT File Chosen!')
    #         tempBool = False
    #
    #     if self.StandardFilePath == "":
    #         self.sg.popup('No Standard File Chosen!')
    #         tempBool = False
    #
    #     if len(self.SelectedStandardsList) == 0:
    #         self.sg.popup('No Standards Selected!')
    #         tempBool = False
    #
    #     if self.LinFilePath != "":
    #         self.LinearityDataAvailable = True
    #
    #     if self.ZeroSetFilePath != "":
    #         self.ZeroSetDataAvailable = True
    #
    #     if tempBool == True:
    #         self.window.FindElement('-dutFile-').Update(disabled=True)
    #         self.window.FindElement('Browse DUT File').Update(disabled=True)
    #         self.window.FindElement('Browse TEGAM File').Update(disabled=True)
    #         self.window.FindElement('-stdFile-').Update(disabled=True)
    #         self.window.FindElement('Select Standards').Update(disabled=True)
    #         self.window.FindElement('-StartCorr-').Update(disabled=True)
    #         self.window.FindElement('-LinBrowseBtn-').Update(disabled=True)
    #         self.window.FindElement('-linFile-').Update(disabled=True)
    #         self.window.FindElement('-SeroSetBrowseBtn-').Update(disabled=True)
    #         self.window.FindElement('-zeroSetFile-').Update(disabled=True)
    #         self.window['-TF1-'].update('')
    #         self.window['-TF2-'].update('(Correction In Process...)', font=('Helvetica', 10, 'bold'))
    #         self.StartCorrectionProcess = True

    def open_window(self):
        # self.thread1 = self.thrd.Thread(target=self._thread_function)
        # self.thread1.start()
        self._thread_function()

    def close_window(self):
        print("attempting to close thread")
        time.sleep(1)
        self.CloseBool = True
        self.window.close()
        # self.thread1.join(timeout=1)
        print("Thread was closed.")


def set_console_size(x=120, y=50):
    import os

    cmd = f'mode {x},{y}'
    os.system(cmd)


def printProgressBar(iteration, total, prefix='', suffix='', decimals=1, length=100, fill='█', printEnd="\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end=printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()


def is_file_empty(filepath):
    # check if file exists
    if os.path.exists(filepath):
        # check if size of file is 0
        if os.stat(filepath).st_size == 0:
            return True
        else:
            return False
    else:
        return True


def DisplayImage_pysimplegui(img_filename, text="", button_name="Continue", xSize=800, ySize=600):
    from PIL import Image
    import io
    import PySimpleGUI as sg

    if os.path.exists(img_filename):
        image = Image.open(img_filename)
        image.thumbnail((ySize, xSize))
        bio = io.BytesIO()
        image.save(bio, format="PNG")

    layout = [
        [sg.Image(key="-IMAGE-", data=bio.getvalue())],
        [
            sg.Text(text),
            # sg.Input(size=(25, 1), key="-FILE-"),
            # sg.FileBrowse(file_types=file_types),

        ],
        [sg.Button(button_name)]
    ]
    window = sg.Window("Image Viewer", layout)

    while True:
        event, values = window.read()
        if event == button_name or event == sg.WIN_CLOSED:
            break

    window.close()


def exercise_step_att_all(exercise_qty=3, interval=0.05):
    printLog("\nExercising Step Attenuators...", console=True)
    for qty in range(1, exercise_qty):
        for i in range(0, 121):
            time.sleep(interval)
            response = step_att_driver(attenuator_resource, i)
    printLog(" > Exercising Step Attenuators Complete!\n", console=True)


def exercise_step_att_step(desired_step, exercise_qty=10, interval=0.1, default_step=0):
    applied_int = interval / 2

    if desired_step == default_step:
        default_step += 1

    for qty in range(1, exercise_qty):
        time.sleep(applied_int)
        response = step_att_driver(attenuator_resource, default_step)
        time.sleep(applied_int)
        response = step_att_driver(attenuator_resource, desired_step)

    time.sleep(applied_int)
    response = step_att_driver(attenuator_resource, desired_step)


class cache:
    @staticmethod
    def put(varname, var_value, variable_cache_file_path="./variable_cache.dat"):
        import os

        def find_line_index(list_var, search_value, split_character="="):
            return_index_found = -1
            for index, line in enumerate(list_var):

                if (search_value in line) and (split_character in line):
                    temp_list = line.split(split_character)
                    line_variable_name = temp_list[0].strip()
                    # input(f"Line var: >{line_variable_name}<")

                    if line_variable_name == search_value:
                        return_index_found = index
                        break

            return return_index_found

        def readTxtFile(filename):
            # Place contents of text files into variable
            f = open(filename, 'r')
            x = f.readlines()
            f.close()
            return x

        def writeListToFile(filename, my_list, write_type='w'):
            def ensure_file_exists(filepath):
                if not os.path.exists(filepath):
                    with open(filepath, 'w') as fp:
                        pass

            def write_list_normal(filename, write_type):
                with open(filename, write_type) as f:
                    for item in my_list:
                        f.write("%s\n" % item)

            def write_list_utf8(filename, write_type):
                with open(filename, write_type, encoding="utf-8") as f:
                    for item in my_list:
                        f.write("%s\n" % item)

            ensure_file_exists(filename)

            try:
                write_list_normal(filename, write_type)
            except Exception as e:
                error = f'{e}'
                error = error.lower()
                if 'permission denied' in error:
                    temp_str = f'Access Denied for file: {filename}\n\n' \
                               f'Please ensure the file is not in use by another program before proceeding!' \
                               f'\nPress Enter to continue...'
                    input(temp_str)
                    write_list_normal(filename, write_type)
                else:
                    write_list_utf8(filename, write_type)

        def write_item_to_file(filename, item_to_write, write_type='a'):
            def ensure_file_exists(filepath):
                if not os.path.exists(filepath):
                    with open(filepath, 'w') as fp:
                        pass

            def write_list_normal(filename, item_to_write, write_type):
                with open(filename, write_type) as f:
                    f.write("%s\n" % item_to_write)

            def write_list_utf8(filename, item_to_write, write_type):
                with open(filename, write_type, encoding="utf-8") as f:
                    f.write("%s\n" % item_to_write)

            ensure_file_exists(filename)

            item_to_write = f"{item_to_write}"

            try:
                write_list_normal(filename, item_to_write, write_type)
            except Exception as e:
                error = f'{e}'
                error = error.lower()
                if 'permission denied' in error:
                    temp_str = f'Access Denied for file: {filename}\n\n' \
                               f'Please ensure the file is not in use by another program before proceeding!' \
                               f'\nPress Enter to Continue... '
                    input(temp_str)
                    write_list_normal(filename, item_to_write, write_type)
                else:
                    write_list_utf8(filename, item_to_write, write_type)

        if not os.path.exists(variable_cache_file_path):
            with open(variable_cache_file_path, 'w') as fp:
                pass

        string_to_write = f"{varname}={var_value}"

        var_cache_contents = readTxtFile(variable_cache_file_path)

        for index, item in enumerate(var_cache_contents):
            item = item.strip()
            var_cache_contents[index] = item

        index_to_write = find_line_index(var_cache_contents, varname)

        if index_to_write < 0:

            write_item_to_file(variable_cache_file_path, string_to_write)
        else:
            var_cache_contents[index_to_write] = string_to_write
            writeListToFile(variable_cache_file_path, var_cache_contents, write_type='w')

    @staticmethod
    def get(varname, variable_cache_file_path="./variable_cache.dat", split_character="="):
        variable_content = ""

        if not os.path.exists(variable_cache_file_path):
            with open(variable_cache_file_path, 'w') as fp:
                pass

        # Open the file
        with open(variable_cache_file_path, "r") as filestream:
            # Loop through each line in the file

            # find_line_index(filestream, "verbose")
            for index, line in enumerate(filestream):

                if (varname in line) and (split_character in line):
                    temp_list = line.split(split_character)
                    line_variable_name = temp_list[0].strip()
                    # input(f"Line var: >{line_variable_name}<")

                    if (line_variable_name == varname) and (len(temp_list) > 1):
                        variable_content = temp_list[1]
                        variable_content = variable_content.strip()
                        break

        filestream.close()
        return variable_content


def sample_power_meter(qty, interval=1, meter_unit="dBm", status_bar=True):
    # Stdev is output as % of nominal

    msmt_list = []
    if debug_flag:
        interval = 0.1
        qty = 3

    if status_bar:
        printLog(f"\n\nSampling Power Meter, {qty} samples", console=True)

    pm_level = queryVisa(pmeter_resource, pmRead)
    # printLog(pm_level)
    pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)

    import time

    # Setup the measurement progress bar
    l = qty
    if status_bar:
        printProgressBar(0, l, prefix='Progress:', suffix='Complete', length=50)

    i = 0
    msmt_list = []
    sdev_List = []
    for number in range(qty):
        time.sleep(interval)
        pm_level = queryVisa(pmeter_resource, pmRead)
        # printLog(pm_level)
        pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)

        # print(f'Msmt {number + 1}: {pm_level} dBm')
        msmt_list.append(pm_level)

        # Sdev values must be stored as linear units
        if meter_unit == "dBm":
            sdev_List.append(10 ** (pm_level / 10))

        if status_bar:
            printProgressBar(i + 1, l, prefix='Progress:', suffix='Complete', length=50)

        i+=1

    avg_msmt = sum(msmt_list) / len(msmt_list)
    if meter_unit == "dBm":
        stdev = statistics.pstdev(sdev_List)
        avg_msmt_for_sdev = sum(sdev_List) / len(sdev_List)
        stdev_pct = stdev / avg_msmt_for_sdev * 100
    else:
        stdev = statistics.pstdev(msmt_list)
        stdev_pct = stdev / avg_msmt * 100

    if status_bar:
        print("\n\n\n")

    return (avg_msmt, stdev_pct)


def initialize_measurement_system():

    def check_if_zero_cal_already():
        temp_str = cache.get("dut_asset")

        if (temp_str == "") or (temp_str != dut_asset):
            return False

        past_ts = cache.get("zero_cal_timestamp")

        if past_ts == "":
            return False

        seconds_in_24_hrs = 60 * 60 * 24

        past_ts = sanitize_variable(past_ts, default_response=0.0, specified_class="float", error_response=0,
                                    eval_operation="at least", eval_threshold=0)

        current_ts = time.time()

        elapsed_seconds = abs(current_ts - past_ts)

        if elapsed_seconds > seconds_in_24_hrs:
            return False
        else:
            return True

    def drift_test():

        perform_test = True

        # Check to see if a result is already saved to the cache file
        temp_str = cache.get("dut_asset")

        if temp_str == dut_asset:
            drift_from_cache = cache.get("drift_test_sdev")

            if drift_from_cache != "":
                drift_from_cache = sanitize_variable(drift_from_cache, default_response=0.0, specified_class="float",
                                                     error_response=999)

                if drift_from_cache != 999:
                    temp_msg = f"Located 30 minute drift test result for this asset:" \
                               f"\n\n{drift_from_cache:.4f}%" \
                               f"\n\nDo you want to re-use this value? (selecting No will re-do the test)"
                    temp_bool = yes_no_popup_simple(temp_msg)

                    if temp_bool:
                        perform_test = False

        if perform_test:
            response = writeVisa(generator_resource, sGenOn)
            test_duration_seconds = 60 * 30
            start_ts = time.time()
            keep_measuring = True
            msmt_bucket = []
            disp_update_cntr = 0
            print("Performing Drift Test...")
            while keep_measuring:
                disp_update_cntr += 1
                current_ts = time.time()
                elapsed_time = abs(current_ts - start_ts)

                if disp_update_cntr > 10:
                    disp_update_cntr = 0
                    seconds_remaining = test_duration_seconds - elapsed_time
                    print(
                        f"> Seconds Remaining: {seconds_remaining:.0f}                                                     \r",
                        end="")

                if elapsed_time > test_duration_seconds:
                    keep_measuring = False

                avg_msmt_dBm, stdev = sample_power_meter(3, interval=0.001, status_bar=False)
                # Values must be stored as linear units, so the percentage math works out properly
                avg_msmt_mW = 10 ** (avg_msmt_dBm / 10)
                msmt_bucket.append(avg_msmt_mW)

            msmt_stdev = statistics.pstdev(msmt_bucket)
            msmt_average = sum(msmt_bucket) / len(msmt_bucket)

            stdev_percent_msd = (msmt_stdev / msmt_average) * 100
            print("> Drift Test Complete                                                                    \n\n\n")
            time.sleep(2)
            response = writeVisa(generator_resource, sGenOff)
        else:
            stdev_percent_msd = drift_from_cache

        cache.put("drift_test_sdev", f"{stdev_percent_msd}")
        return stdev_percent_msd
    # debug_flag = False
    # User Instructions

    # =========================================================
    #                    PM Setup Routine
    # =========================================================
    if debug_flag == False:

        if check_if_zero_cal_already() == True:
            temp_msg = "The DUT sensor appears to have been zeroed and calibrated already." \
                       "\n\nDo you want to re-Zero/Cal the sensor?"
            temp_bool = yes_no_popup_simple(temp_msg)
        else:
            temp_bool = True

        if temp_bool:
            temp_msg = f'- Connect the DUT sensor to the power meter channel 1/A\n\n' \
                       f'- Connect the sensor to the {pMeterName} Reference Port\n\n'
            DisplayImage_pysimplegui("./Images/Zero_Cal_Connection.png", text=temp_msg, button_name="Continue", xSize=1600,
                                     ySize=1600)
            temp_msg = "Do not continue until the sensor has warmed-up for at least 30 minutes." \
                       "\n\nNote: Failure to follow warmup time will likely result in the sensor" \
                       " failing (this was particularly noted with E series sensors)."
            msg_box_simple(temp_msg)

            temp_msg = "Please wait for the zero/calibration process to complete..."
            msg_box_simple(temp_msg)

            # Zero and Cal the RF Power Meter
            if debug_flag == False:
                printLog("\nZeroing the power sensor...")
                printLog("Reset Command", console=verbose_flag)
                response = writeVisa(pmeter_resource, pmRst, opc=True, response="1")
                printLog(response, console=verbose_flag)
                printLog(f"Frequency set to {test_frequency} Hz...", console=verbose_flag)
                response = writeVisa(pmeter_resource, pmFreq.replace('<val>', test_frequency))
                printLog(response, console=verbose_flag)
                printLog("Send pmdBmeas command ...", console=verbose_flag)
                response = writeVisa(pmeter_resource, pmdBmeas, opc=True)
                printLog(response, console=verbose_flag)
                printLog("Send pmZero command ...", console=verbose_flag)
                response = writeVisa(pmeter_resource, pmZero, opc=True)
                printLog("Getting zero response...", console=verbose_flag)
                printLog(response, console=verbose_flag)
                printLog("> Zeroing complete.\n")
                printLog("Reference calibrating the power sensor...")
                response = writeVisa(pmeter_resource, pmCal, opc=True)
                printLog(response, console=verbose_flag)
                printLog("> Reference cal complete.\n")
                cache.put("dut_asset", dut_asset)
                cache.put("zero_cal_timestamp", f"{time.time()}")
        else:
            temp_msg = '- Ensure the DUT sensor is connected to power meter channel 1/A.'
            msg_box_simple(temp_msg)

            response = writeVisa(pmeter_resource, pmRst, opc=True)
            printLog(response, console=verbose_flag)
            response = writeVisa(pmeter_resource, pmFreq.replace('<val>', test_frequency))
            printLog(response, console=verbose_flag)

    # =========================================================
    #            Measurement Assembly Setup Routine
    # =========================================================
    if debug_flag == False:
        temp_msg = f'- Connect the {step_att_name} channel X to the 11 dB Step Attenuator\n\n' \
                   f'- Connect the {step_att_name} channel Y to the 110 dB Step Attenuator\n\n' \
                   f'- Connect the {generator_name} generator OUTPUT to the {step_att_name} attenuator stack INPUT\n\n' \
                   f'- Connect the {step_att_name} attenuator stack OUTPUT to the DUT INPUT\n\n' \
                   f'- Wait for the measurement process to complete...'
        DisplayImage_pysimplegui("./Images/Basic_Connection.png", text=temp_msg, button_name="Continue", xSize=1600,
                                 ySize=1600)
        # msg_box_simple(temp_msg)


    # =========================================================
    #            Exercise The Step Attenuator
    # =========================================================
    if debug_flag == False and exercise_att == True:
        exercise_step_att_all(exercise_qty=1)

    # =========================================================
    #                    Sig Gen Setup Routine
    # =========================================================
    if debug_flag == False:
        printLog("Setting Sig Gen...")
        response = writeVisa(generator_resource, sGenRst)
        printLog(response, console=verbose_flag)
        response = writeVisa(generator_resource, sGenPowSet.replace('<val>', '-100'))
        printLog(response, console=verbose_flag)
        response = writeVisa(generator_resource, sGenFreqSet.replace('<val>', test_frequency))
        printLog(response, console=verbose_flag)
        printLog("> Sig Gen setup complete.\n")

        # Get the max linearity step
        printLog("\nChecking if Sig Gen can achieve max output level...")
        gen_power = lin_steps_list[-1]

        # Attempt to level the generator to the max step through all the cables etc.
        response = step_att_driver(attenuator_resource, 0)
        printLog(response, console=verbose_flag)

        # output_limit =  mW_dBm(dBm_mW(gen_power) * 1.5)
        output_limit =  mW_dBm(dBm_mW(gen_power) * 1)
        level_generator_and_power_meter(gen_power, leveling_tol=0.10, settling_time=5, safe=True, max_output=output_limit)

        # Get the power level setting that was achieved during the generator leveling process
        gen_level_pow = queryVisa(generator_resource, sGenPowRead)
        gen_level_pow = sanitize_variable(gen_level_pow, specified_class='float',
                                                  default_response=999_999_999)

        gen_level_pow_mW = dBm_mW(gen_level_pow)
        gen_level_ll_mW = dBm_mW(gen_level_pow) * 0.5
        gen_level_ul_mW = dBm_mW(gen_level_pow) * 1.5

        if gen_level_pow_mW < gen_level_ll_mW or gen_level_pow_mW > gen_level_ul_mW:
            temp_msg = f"Leveled power set at the DUT is too high or too low!" \
                       f"\n\nSet: {gen_level_pow:.4f} dBm" \
                       f"\nTarget: {gen_power:.4f} dBm"

            error_and_exit(messag=temp_msg)

        # Assign the set power level to a global variable for use by the linearity msmt routine
        global gen_level_pow_for_lin_msmt
        gen_level_pow_for_lin_msmt = gen_level_pow

        # Set the generator to the max linearity step value
        retry_bool = True
        while retry_bool:
            check_leveled = True


            response = writeVisa(generator_resource, sGenFreqSet.replace('<val>', test_frequency))
            printLog(response, console=verbose_flag)



            response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'{gen_level_pow_for_lin_msmt}'))
            printLog(response, console=verbose_flag)

            # Check if the generator assumed the specified power level
            response = queryVisa(generator_resource, sGenPowRead)
            response = sanitize_variable(response, specified_class='float',
                                              default_response=999_999_999)
            try:
                response = float(response)
                if not response == gen_level_pow_for_lin_msmt:
                    check_leveled = False
                    temp_str = f'Cannot set the signal generator to the specified output level!\n\n' \
                               f'Specified level: {gen_level_pow_for_lin_msmt}, Generator Max: {response}\n\n' \
                               f'Do you want to retry?'
                    temp_bool = yes_no_popup_simple(temp_str)

                    if temp_bool == False:
                        temp_str = f"Could not set the generator output level to {gen_level_pow_for_lin_msmt} dBm!"
                        error_and_exit(temp_str)
                else:
                    check_leveled = True

            except:
                check_leveled = False
                temp_str = f'Cannot set the signal generator to the specified output level!\n\n' \
                           f'Specified level: {gen_level_pow_for_lin_msmt}, Generator Max: {response}\n\n' \
                           f'Do you want to retry?'
                temp_bool = yes_no_popup_simple(temp_str)

                if temp_bool == False:
                    temp_str = f"Could not set the generator output level to {gen_level_pow_for_lin_msmt} dBm!"
                    error_and_exit(temp_str)

            if check_leveled:
                response = writeVisa(generator_resource, sGenOn)
                printLog(response, console=verbose_flag)

                time.sleep(10)
                printLog('Checking if generator is unleveled...', console=verbose_flag)
                response = queryVisa(generator_resource, unlevel_err_check)
                printLog(f'Sig gen leveled query response: {response}', console=verbose_flag)
                if response == unlevel_err_response:
                    temp_str = f'Cannot achieve leveled output power on the signal generator!\n\n' \
                               f'Desired level: {gen_level_pow_for_lin_msmt}\n\n' \
                               f'Do you want to retry?'
                    temp_bool = yes_no_popup_simple(temp_str)

                    if temp_bool == False:
                        temp_str = "Could not level the generator output!"
                        error_and_exit(temp_str)
                else:
                    retry_bool = False

            response = writeVisa(generator_resource, sGenOff)
            printLog(response, console=verbose_flag)
        printLog("> Sig Gen leveled output check complete.\n")

        # =========================================================
        #                    Drift Test Routine
        # =========================================================
        global drift_test_bool
        if drift_test_bool:
            temp_msg = "Do you want to run the system drift test?" \
                       "\n\nNote: this takes 30 minutes"
            drift_test_bool = yes_no_popup_simple(temp_msg)

        if drift_test_bool:
            global drift_test_result_pct
            drift_test_result_pct = drift_test()

def perform_lin_msmt(steps_list, tol_list, ref_pow, sample_qty):
    def dut_bias_msmt(sample_qty, reference_pow_setting, gen_pow_level, msmt_interval=1, settle_time=10):
        print("Performing DUT bias measurement...")
        if debug_flag:
            settle_time = 1

        # Turn on the signal generator
        response = writeVisa(generator_resource, sGenOn)
        printLog(response, console=verbose_flag)
        response = writeVisa(generator_resource, sGenFreqSet.replace('<val>', test_frequency))
        printLog(response, console=verbose_flag)

        # Set the step ATT
        exercise_step_att_step(int(lin_step_to_att_set_dict[reference_pow_setting]), exercise_qty=20)
        log.info(f'reference_pow_setting for bias msmt: {reference_pow_setting}')
        log.info(f'gen_pow_level for bias msmt: {gen_pow_level}')
        log.info(f'attenuator setting for bias msmt: {int(lin_step_to_att_set_dict[reference_pow_setting])}')
        response = step_att_driver(attenuator_resource, int(lin_step_to_att_set_dict[reference_pow_setting]))
        printLog(response, console=verbose_flag)

        response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'{gen_pow_level}'))
        printLog(response, console=verbose_flag)

        pm_level = queryVisa(pmeter_resource, pmRead)
        printLog(pm_level, console=verbose_flag)
        pm_level = sanitize_variable(pm_level, specified_class='float', default_response=999_999_999)

        print(f'Settling measurement system for {settle_time} seconds...')
        time.sleep(settle_time)

        avg_msmt, stdev_pct = sample_power_meter(sample_qty, interval=msmt_interval)
        log.info(f'dut power read for bias msmt: {avg_msmt}')

        avg_msmt = reference_pow_setting - avg_msmt

        log.info(f'final bias_msmt value: {avg_msmt}, stdev_pct: {stdev_pct}')

        print(" > DUT bias measurement complete!")

        return (avg_msmt, stdev_pct)

    def att_debug_test_routine(lin_step_to_att_set_dict):
        loop = True
        print("======================================")
        print("    Attenuator Debug Test Routine")
        print("======================================")

        while loop:
            try:
                step = input("Enter a desired linearity step (\"999\" exits): ")
                step = step.strip()
                step = int(step)
                step = abs(step) * -1

                if step == -999:
                    loop = False
                else:
                    true_att = lin_step_to_att_set_dict[step]
                    attenuation_value = lin_step_att_actual[step]
                    response = step_att_driver(attenuator_resource, true_att)
                    temp_msg = f"response\n" \
                               f"Attenuation Value: {attenuation_value} dBm"
                    print(temp_msg)
            except Exception as error:
                print(f"{error}")

    def msr_lin_steps(settle_time=5, sample_qty=1):
        sample_qty_if_failed_step = 31
        if debug_flag:
            settle_time = 0.5
            sample_qty = 1

        # Zero the power sensor
        # if not debug_flag:
        #     print("Zeroing the power meter...")
        #     response = writeVisa(generator_resource, sGenOff)
        #     printLog(response, console=verbose_flag)
        #     response = writeVisa(pmeter_resource, pmZero, opc=True)
        #     printLog(response, console=verbose_flag)

        # Turn on the signal generator and set the power level
        # gen_level_pow_for_lin_msmt is a global variable set during the generator setup process
        response = writeVisa(generator_resource, sGenPowSet.replace('<val>', f'gen_level_pow_for_lin_msmt'))
        printLog(response, console=verbose_flag)
        response = writeVisa(generator_resource, sGenOn)
        printLog(response, console=verbose_flag)

        lin_steps_msmt = []
        lin_steps_msmt_mW = []
        lin_steps_dev = []
        lin_steps_nominal_mW = []
        lin_steps_lower_mW = []
        lin_steps_upper_mW = []
        lin_step_unc = []
        eval_str_list = []
        fail_bool_list = []
        att_ref_step_setting = int(lin_step_to_att_set_dict[ref_pow])
        overall_fail_flag = False
        columnar_data_list = []
        for index, lin_step in enumerate(steps_list):
            loop_step = True
            loop_step_cntr = 0

            while loop_step:

                loop_step_cntr+=1
                if loop_step_cntr > 1:

                    printLog(f"Last msmt attempt failed? {fail_bool}. Eval Type: {eval_str}\n Retrying the measurement...", console=verbose_flag)
                    response = step_att_driver(attenuator_resource, 110)

                    printLog(response, console=verbose_flag)
                    printLog(f"Settling Sensor {settle_time} seconds", console=verbose_flag)
                    time.sleep(settle_time)
                    avg_msmt, stdev = sample_power_meter(3, interval=sampling_intv)

                    printLog(f"DUT measured: {avg_msmt} dBm, at no power state", console=verbose_flag)

                    exercise_step_att_step(int(lin_step_to_att_set_dict[lin_step]))

                if verbose_flag:
                    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                printLog("\n++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++", console=verbose_flag)

                # Turn on the signal generator

                response = writeVisa(generator_resource, sGenOn)
                printLog(response, console=verbose_flag)

                response = writeVisa(pmeter_resource, pmFreq.replace('<val>', test_frequency))
                printLog(response, console=verbose_flag)

                response = writeVisa(generator_resource, sGenFreqSet.replace('<val>', test_frequency))
                printLog(response, console=verbose_flag)

                # Get step attenuator data:
                true_att = lin_step_att_actual[lin_step]
                att_unc = lin_step_unc_actual[lin_step]

                # Allow testing of the attenuator, if desired
                test_att_debug = False
                if test_att_debug == True or (dut_asset == "att_test"):
                    att_debug_test_routine(lin_step_to_att_set_dict)

                # Set the att to the current linearity step
                exercise_step_att_step(int(lin_step_to_att_set_dict[lin_step]), exercise_qty=3)
                response = step_att_driver(attenuator_resource, int(lin_step_to_att_set_dict[lin_step]))
                response = step_att_driver(attenuator_resource, int(lin_step_to_att_set_dict[lin_step]))
                printLog(response, console=verbose_flag)

                # Allow the measurement to settle
                printLog(f'Settling linearity step: {lin_step} dBm (actual: {true_att:.4f} dBm)')
                time.sleep(settle_time)

                # get the current step measurement
                if loop_step_cntr > 1:
                    sample_qty_current = sample_qty_if_failed_step
                else:
                    sample_qty_current = sample_qty
                avg_msmt, stdev_pct = sample_power_meter(sample_qty_current, interval=sampling_intv)
                avg_msmt_normalized = avg_msmt + bias_msmt
                printLog(
                    f"DUT measured: {avg_msmt} dBm before normalization, {avg_msmt_normalized} dBm after normalization...",
                    console=verbose_flag)

                printLog(f"Perform Linearity Step Normalization? {normalize_flag}")
                if normalize_flag:
                    printLog(
                        f"DUT measured: {avg_msmt_normalized} dBm before Lin Step Normalization...",
                        console=verbose_flag)
                    power_normalization_delta = lin_step - true_att
                    avg_msmt_normalized += power_normalization_delta
                    printLog(
                        f"DUT measured: {avg_msmt_normalized} dBm after Lin Step Normalization...",
                        console=verbose_flag)

                    # Get the tolerances of the current step
                    tolerance = tol_list[index]
                    true_att = true_att + power_normalization_delta
                    lin_step_mW = dBm_mW(true_att)
                    lin_steps_nominal_mW.append(lin_step_mW)
                    upper_lim_mW = lin_step_mW + (lin_step_mW / 100 * tolerance)
                    lin_steps_upper_mW.append(upper_lim_mW)
                    lower_lim_mW = lin_step_mW - (lin_step_mW / 100 * tolerance)
                    lin_steps_lower_mW.append(lower_lim_mW)

                else:
                    # Get the tolerances of the current step
                    tolerance = tol_list[index]
                    lin_step_mW = dBm_mW(true_att)
                    lin_steps_nominal_mW.append(lin_step_mW)
                    upper_lim_mW = lin_step_mW + (lin_step_mW / 100 * tolerance)
                    lin_steps_upper_mW.append(upper_lim_mW)
                    lower_lim_mW = lin_step_mW - (lin_step_mW / 100 * tolerance)
                    lin_steps_lower_mW.append(lower_lim_mW)

                avg_msmt_mW = dBm_mW(avg_msmt_normalized)
                lin_steps_msmt_mW.append(avg_msmt_mW)

                # Calculate the percent of tolerence
                percent_tol = abs(((lin_step_mW - avg_msmt_mW) / (lin_step_mW / 100 * tolerance)) * 100)

                lin_steps_msmt.append(avg_msmt_normalized)
                lin_steps_dev.append(stdev_pct)

                # Convert all uncertainty contributors to %
                type_a_dB = percent_to_dB(stdev_pct)

                bias_stdev_dB = percent_to_dB(bias_stdev_pct)

                att_unc_percent = dBtoPercent(att_unc)

                temp_str = f'Uncertainty Contributors: ' \
                           f'Drift test (1 sdev): {drift_test_result_pct:.2f}%, ' \
                           f'type_a: {type_a_dB:.4f} dB ({stdev_pct:.4f}%), ' \
                           f'bias_stdev: {bias_stdev_dB:.4f} dB ({bias_stdev_pct:.4f}%), ' \
                           f'att_unc: {att_unc:.4f} dB ({att_unc_percent:.4f}%)'
                printLog(temp_str, console=verbose_flag)

                # Calculate measurement uncertainty

                unc_2s_mW = calc_uncertainty(lin_step_mW, stdev_pct, drift_test_result_pct, sample_qty_current, attenuator_unc_pct=att_unc_percent,
                                             bias_sdev_pct=bias_stdev_pct)
                lin_step_unc.append(unc_2s_mW)

                # Calculate the TUR
                tur = abs((lin_step_mW / 100 * tolerance) / unc_2s_mW)

                # Evaluate the measurement
                eval_str, fail_bool = Pass_Fail_Eval(avg_msmt_mW, lower_lim_mW, upper_lim_mW, unc_2s_mW)
                eval_str_list.append(eval_str)
                fail_bool_list.append(fail_bool)

                # Set the overall fail flag
                if fail_bool == True:
                    if loop_step_cntr > 1:
                        fail_bool_list[index] = fail_bool
                        loop_step = False
                elif (eval_str.lower() == 'ugb1') and (loop_step_cntr < 2) and (auto_retest == True):
                    loop_step = True
                else:
                    loop_step = False

                # Create measurement line
                # db,LL,mW,UL,unc,eval,fail_bool
                # lin_step = int(lin_step)

                msmt_file_header = 'lin_step,att_actual,msmt(dB),lower_lim(mW),msmt(mW),upper_lim(mW),unc_2s(mW),eval,fail_bool,percent_tol,TUR,dut_tol(percent)'
                tofile =          f'{lin_step},{true_att},{avg_msmt_normalized},{lower_lim_mW},{avg_msmt_mW},{upper_lim_mW},{unc_2s_mW},{eval_str},{fail_bool},{percent_tol},{tur},{tolerance}'

                # Determine how much resolution to apply
                res_qty = calc_resol_qty(avg_msmt_mW, additonal_res=4)
                if verbose_flag:
                    print(f'res_qty: {res_qty}')

                # Setup and display columnar for measurement display
                headers = ['Step (dBm)', 'Actual (dBm)', 'Measured (dBm)', 'Lower Limit (mW)', "Measured (mW)", "Upper Limit (mW)", "Uncertainty (mW)", "Evaluation", "Step Fails?", "% of Tol.", "TUR"]
                data = [f'{lin_step:.0f}', f'{true_att:.4f}', f'{avg_msmt_normalized:.4f}', f'{lower_lim_mW:.{res_qty}f}',
                     f'{avg_msmt_mW:.{res_qty}f}', f'{upper_lim_mW:.{res_qty}f}', f'{unc_2s_mW:.1e}', f'{eval_str}',
                     f'{fail_bool}', f'{percent_tol:.1f} %', f'{tur:.1f}:1', ]


                console_text_width = 1000
                columnar_data_list.append(data)
                table = columnar(columnar_data_list, headers, no_borders=False, min_column_width=20, terminal_width=console_text_width)

                # Ensure the console is the correct size to display the data
                set_console_size(x=console_text_width)

                print(table)

            # file_path = 'mysample.txt'

            # Check if the measurement file is still empty; write csv header if so
            if is_file_empty(msmt_file_name) == True:
                writeListToFile(msmt_file_name, [msmt_file_header], write_type='a')

            # Write measurement to file
            writeListToFile(msmt_file_name, [tofile], write_type='a')
            if debug_flag:
                input("::: ===========================================================================================")

            # Check to see if any steps failed and set the test flag accordingly
            for item in fail_bool_list:
                if item == True:
                    overall_fail_flag = True


        return overall_fail_flag

    # Build the step setting dictionary (key = power level, value = att setting)
    lin_step_to_att_set_dict = build_step_setting_dict(steps_list, refStepSetting)
    log.info('lin_step_to_att_set_dict:')
    log.info(lin_step_to_att_set_dict)

    # Build the true attenuation value for each step, according to the attenutor data, relative to the reference step
    # Build the true unc value for each step, according to the attenutor data, relative to the reference step
    lin_step_att_actual, lin_step_unc_actual = build_step_nominal_dict(lin_step_to_att_set_dict, refStepSetting)
    log.info('lin_step_unc_actual:')
    log.info(lin_step_unc_actual)
    log.info('lin_step_att_actual:')
    log.info(lin_step_att_actual)

    # Obtain the sensor bias msmt value
    true_att = lin_step_att_actual[refStepSetting]
    att_ref_step_setting = int(lin_step_to_att_set_dict[ref_pow])
    exercise_step_att_step(att_ref_step_setting, exercise_qty=3)
    response = step_att_driver(attenuator_resource, att_ref_step_setting)
    bias_msmt, bias_stdev_pct = dut_bias_msmt(biasMsmtQty, true_att, gen_level_pow_for_lin_msmt, msmt_interval=sampling_intv)

    if verbose_flag:
        print(f'bias_msmt {bias_msmt}, {bias_stdev_pct}')
    if debug_flag:
        input(':::')

    dut_failed = msr_lin_steps(settle_time=settlingTime, sample_qty=samplingQuantity)

    return dut_failed



#ToDo

# Don't save a result during testing if it is worse than the one prior (will need to create a results buffer)
# When clicking "no" to confirm the VISA resource, change it so that you can manually select (currently retries)
# Check the sig gen max PLA routine to ensure it is done properly
# Require sensor zero and cal
# Level generator to the max required level to include cable and connections
# Fix crash if user exits list screen when selecting a VISA resource
# Clean-up console measurement display

# Start Program =================================================================

# Set initial program variables
logFile = 'SensorLinCal.log'
log.basicConfig(filename=logFile, filemode='a', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S', level=log.INFO)

printLog('--------------- New program instance ---------------')
cwd = os.getcwd() + "\\"
clear()
userInterfaceHeader(program_name, cs_number, program_version, cwd, logFile)
print("=======================================================================")

# Pull in settings from the config file ------------
configFile = "Configuration.cfg"
get_config_file_settings()

# Create the measurement templates folder, if it does not already exist`
path_exist, error_info = check_and_create_path(msmt_templates_folder)
if not path_exist:
    tempStr = f'Could not locate msmt_results_folder path: {msmt_templates_folder}\n' \
              f'Attempted to create this path without success, got error:\n' \
              f'{error_info}' \
              f'Please fix this path issue and re-run the program'
    error_and_exit(messag=tempStr)

# Create the measurement results folder, if it does not already exist
path_exist, error_info = check_and_create_path(msmt_results_folder)
if not path_exist:
    tempStr = f'Could not locate msmt_results_folder path: {msmt_results_folder}\n' \
              f'Attempted to create this path without success, got error:\n' \
              f'{error_info}' \
              f'Please fix this path issue and re-run the program'
    error_and_exit(messag=tempStr)



# Check that all set files exist
temp_list = [PS_CalResultsFolder, linBudgetTxtFile, generator_driver, pm_driver, attenuator_driver]
verify_file_paths(temp_list)

# Check that the budget file is up to date
check_unc_budget(linBudgetTxtFile)
printLog('Linearity budget file is valid...')

# Initialize the GUI Interface
program = GuiProgramWindow(f"Power Sensor Linearity Calibration, v. {program_version} - The Cal Lab Is On Fire")

# Set GUI Window Variables
program.att_data_initial_path11 = linearityCalDataFilePath11
program.att_data_initial_path110 = linearityCalDataFilePath110
program.DutDataFolder = msmt_templates_folder
program.StdDataFolder = standardsDataFolder

# Create the GUI window
program.open_window()

# Get the attenuator standards data
att_data_list11 = program.att_data_list11
att_data_list110 = program.att_data_list110

# Check the auto-retest flag
auto_retest = program.retest

# Check the drift_test flag and set the initial value to zero
drift_test_bool = program.drift_test
drift_test_result_pct = 0.0

# Set Visa Resources
attenuator_resource = program.set_attenuator_resource
generator_resource = program.set_generator_resource
pmeter_resource = program.set_pmeter_resource

# Get DUT identification details
dut_asset = program.dut_asset_number
cur_dt = get_convert_timestamp()
msmt_file_name = f'{msmt_results_folder}/Lin Data - {dut_asset} - {cur_dt}.csv'
plot_file_name = f'{msmt_results_folder}/Lin Data - {dut_asset} - {cur_dt}.png'

msmt_file_name = standardize_file_path_format(msmt_file_name)
plot_file_name = standardize_file_path_format(plot_file_name)

# Check if measurement file exists
msmt_file_name = rename_if_file_exists(msmt_file_name)
plot_file_name = rename_if_file_exists(plot_file_name)

# Perform the DUT measurement process:
run_again = True
dut_failed = False
while run_again:
    initialize_measurement_system()
    dut_failed = perform_lin_msmt(lin_steps_list, lin_steps_tol, refStepSetting, samplingQuantity)

    if dut_failed:
        temp_str = 'DUT failed the linearity measurement!\n\n Retry measurement?'
        run_again = yes_no_popup_simple(temp_str)
    else:
        run_again = False

    if run_again:
        # Pull in settings from the config file ------------
        configFile = "Configuration.cfg"
        get_config_file_settings()

        # Check that all set files exist
        temp_list = [msmt_results_folder, PS_CalResultsFolder, linBudgetTxtFile, generator_driver, pm_driver, attenuator_driver]
        verify_file_paths(temp_list)

        # Check that the budget file is up to date
        check_unc_budget(linBudgetTxtFile)
        printLog('Linearity budget file is valid...', console=verbose_flag)

        # Prompt User For DUT PS-Cal File, and import file path and test parameters
        get_dut_template_data()

        # Get the attenuator standards data
        att_data_list11, att_data_list110 = get_attenuator_standard_data(linearityCalDataFilePath11,
                                                                         linearityCalDataFilePath110)

        # Check if measurement file exists
        msmt_file_name = rename_if_file_exists(msmt_file_name)

# Reset Instruments
response = writeVisa(generator_resource, sGenRst)
printLog(response, console=verbose_flag)
response = writeVisa(pmeter_resource, pmRst)
printLog(response, console=verbose_flag)

# Create a PS-Cal Corrector Compatible linearity data file
printLog("Creating PS-Cal Corrector compatible linearity file...")
dat_filepath, ps_cal_file_path = create_pscalcorr_lin_dat_file(msmt_file_name, PS_CalResultsFolder)
shutil.copy(dat_filepath, ps_cal_file_path)
printLog("Successfully created PS-Cal Corrector compatible linearity file")

# Plot the test results
try:
    printLog("Setting up data plot to save", console=verbose_flag)
    plot_filepath = plot_data(msmt_file_name, plot_file_name, plt_x=plot_x_inches, plt_y=plot_y_inches)
    printLog("Data plot was saved", console=verbose_flag)
except Exception as e:
    temp_msg = f"Attempted to plot the calibration output graph, but got this error:" \
               f"\n\n{e}" \
               f"\n\n Plotting data is not required to pass the test, at this time, as the " \
               f"\nmeasurement data is saved to the CSV file." \
               f"\n\nIf you like to see the pretty plots then make sure the x and y size in inches" \
               f"\nin the program configuration file is not set to something ridiculous."
    msg_box_simple(temp_msg)

# Finalize the program
printLog("Finalizing program", console=verbose_flag)
eval_str = 'FAIL' if dut_failed else 'PASS'
temp_str = f'CALIBRATION COMPLETE! - DUT {eval_str}\n\n' \
           f'EVALUATION: {eval_str}\n\n' \
           f'PS-Cal Corrector Import File: {dat_filepath}\n\n' \
           f'CSV Results: {msmt_file_name}'
DisplayImage_pysimplegui(plot_filepath, text=temp_str, button_name="Continue", xSize = 2880, ySize = 1620)

printLog("Program complete.")
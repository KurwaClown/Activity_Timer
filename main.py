import sys
from win32gui import GetWindowText, GetForegroundWindow
import win32process
import time
import datetime
import uiautomation as auto
import json
import os
import asyncio
import wmi
import threading

lastWindow = ""
lastUrl = ""
siteList = []
activityList = []

def get_active_window():
    """
    Get active window name
    """
    return GetWindowText(GetForegroundWindow())


def is_new_window(newWindow, storedWindow):
    """
    Check if the current active window is different
    """
    if newWindow != storedWindow:
        return True
    else:
        return False

def get_browser_url():
        window = GetWindowText(GetForegroundWindow())
        browserControl = auto.Control(searchDepth =1 , Name=str(window))
        edit = browserControl.DocumentControl()
        try:
            return str(edit.GetValuePattern().Value)
        except LookupError as e:
            return None
        
def time_converter(seconds):
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    return {"seconds" : seconds, "minutes" : minutes, "hours" : hours}


def dataDump():
    threading.Timer(15, dataDump).start()
    with open("Activities.json", "w") as json_file:
        json.dump(activities, json_file, indent=4, ensure_ascii=False)
    



try:
    file = open("Activities.json", "r")
    activities = json.loads(file.read())
    for activity in activities["activities"]:
        activityList.append(activity['name'])
    for site in activities['sites']:
        siteList.append(site['url'])
except FileNotFoundError as e:
    file = open("Activities.json", "w")
    activities = {"activities" : [], "sites" : []}
    json.dump(activities, file, indent=4)
finally:
    file.close()
dataDump()
while True:
    try: 
        startTime = datetime.datetime.now()
        while True:
            newWindow = get_active_window()

            pid = win32process.GetWindowThreadProcessId(GetForegroundWindow())[1]
            pidInfo = os.popen(f'TASKLIST /FI "PID eq {pid}" /FO CSV').read()
            pidExe = pidInfo[85:].replace('"','').split(',')[0]

            procs = wmi.WMI().Win32_Process(name=pidExe)
            for proc in procs:
                appPath = proc.ExecutablePath
                break
            
            newWindowName = appPath.split('\\')[-2]
                
            if is_new_window(newWindowName, lastWindow) and newWindowName!= "":
                lastWindow = newWindowName

            if appPath.split('\\')[-1] in ["opera.exe", "Google Chrome"]:
                try:
                    lastUrl = get_browser_url().split("/")[2]
                    if lastUrl.startswith("www"):
                        lastUrl = lastUrl[4:]
                    

                    if lastUrl not in siteList:
                        datas = {
                            "url" : str(lastUrl),
                            "first_use" : str(time.strftime("%d %b %Y %H:%M:%S", time.localtime(time.time()))),
                            "time_used" : {
                                "days" : 0,
                                "hours" : 0,
                                "minutes" : 0,
                                "seconds" : 0,
                                "microseconds": 0,
                            },
                            "last_use" : str(time.strftime("%d %b %Y %H:%M:%S", time.localtime(time.time())))
                        }
                        siteList.append(lastUrl)
                        activities['sites'].append(datas)
                    else:
                        currentActivity = activities["sites"][siteList.index(str(lastUrl))]
                        timeDict = {"microseconds": 1000000, "seconds": 60, "minutes": 60, "hours":24}
                        timeUsed = currentActivity['time_used']
                        currentActivity['last_use'] = str(time.strftime("%d %b %Y %H:%M:%S", time.localtime(time.time())))
                        endTime = datetime.datetime.now()

                        deltaTime = endTime - startTime
                        convertedTime = time_converter(deltaTime.seconds)

                        extraTime = 0
                        for timeValue in timeDict:
                            
                            if timeValue=="microseconds":
                                timeUsed[timeValue] += int((deltaTime.microseconds))
                            else:
                                timeUsed[timeValue] += convertedTime[timeValue]+extraTime
                            
                            extraTime = int(timeUsed[timeValue]/timeDict[timeValue])
                            timeUsed[timeValue] = int(timeUsed[timeValue]%timeDict[timeValue])

                except AttributeError as e:
                    print("Error getting site url : no URL retrievable")


            if lastWindow not in activityList:
                datas = {
                    "name" : str(lastWindow),
                    "first_use" : str(time.strftime("%d %b %Y %H:%M:%S", time.localtime(time.time()))),
                    "time_used" : {
                        "days" : 0,
                        "hours" : 0,
                        "minutes" : 0,
                        "seconds" : 0,
                        "microseconds": 0,
                    },
                    "last_use" : str(time.strftime("%d %b %Y %H:%M:%S", time.localtime(time.time()))),
                    "exe_path": ""
                    }

                activityList.append(lastWindow)
                activities['activities'].append(datas)
            else:
                currentActivity = activities["activities"][activityList.index(str(lastWindow))]
                timeDict = {"microseconds": 1000000, "seconds": 60, "minutes": 60, "hours":24}
                timeUsed = currentActivity['time_used']
                currentActivity['last_use'] = str(time.strftime("%d %b %Y %H:%M:%S", time.localtime(time.time())))
                endTime = datetime.datetime.now()

                deltaTime = endTime - startTime
                convertedTime = time_converter(deltaTime.seconds)

                currentActivity['exe_path'] = appPath

                extraTime = 0
                for timeValue in timeDict:
                    
                    if timeValue=="microseconds":
                        timeUsed[timeValue] += int((deltaTime.microseconds))
                    else:
                        timeUsed[timeValue] += convertedTime[timeValue]+extraTime
                    
                    extraTime = int(timeUsed[timeValue]/timeDict[timeValue])
                    timeUsed[timeValue] = int(timeUsed[timeValue]%timeDict[timeValue])
            print(activities)
            startTime = datetime.datetime.now()
            time.sleep(1)
    except Exception as e:
        with open ("error.logs", "a") as error_log:
            error_log.write(f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} : {e}\n")
        with open ("Activities.json", "w") as file:
            json.dump(activities, file, indent=4, ensure_ascii=False)

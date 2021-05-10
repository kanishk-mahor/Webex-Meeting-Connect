# ======================================================================================================================
#                 Copyright (C) 2020 Kanishk Mahor - All rights reserved
# ========================================================================================
# Notice:  All Rights Reserved.
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# ======================================================================================================================

import datetime
import warnings
import time
import re
import sched
import win32com.client

# --------------------
# pip install selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

# ---------------------
# pip install pywinauto
from pywinauto.keyboard import send_keys
from pywinauto.application import Application


# ------------------------------------------------------------------------------------------------
# Chrome webbrowser driver path :you can download from https://chromedriver.chromium.org/downloads
PATH = "C:/Program Files/chromedriver.exe"

# --------------------------
# creating schedule instance
schedule = sched.scheduler(time.time, time.sleep)

# ----------------------------------
# Outlook calender API to get events


def get_calender():
    outlook = win32com.client.Dispatch(
        'Outlook.Application').GetNamespace('MAPI')
    calender = outlook.getDefaultFolder(9).Items
    # Including recurring events
    calender.IncludeRecurrences = True
    calender.Sort('[Start]')
    # ----------today date-----------
    today = datetime.datetime.today()
    begin = today.date().strftime("%d/%m/%Y")
    # -------tomorrow date from today----------
    tomorrow = datetime.timedelta(days=1)+today
    end = tomorrow.date().strftime("%d/%m/%Y")
    # -------------restrict calender events to today only ---------------
    restriction = "[Start] >= '" + begin + "' AND [END] <= '" + end + "'"
    calender = calender.Restrict(restriction)
    events = {'Start': [], 'Subject': [], 'Body': []}
    for a in calender:
        events['Start'].append((a.start).strftime("%H:%M"))
        events['Subject'].append(a.Subject)
        events['Body'].append(a.body)
    return events

# ----------------------------
# join metting at metting time


def join(calender_return, current_time):

    # ----List if all todays meeting-----
    meet = list(calender_return['Start'])
    # ----index of current meeting----
    to_join = meet.index(current_time)
    # -extracting body content of current meeting-
    link1 = list(calender_return['Body'])[to_join]
    # -------------------------Parsing url from body-----------------------
    link_to_go = re.search("(?P<url>https?://[^\\s]+)", link1).group("url")
    link_to_go = link_to_go[:-1]

    # wait for one minute before joing meeting
    time.sleep(60)

    # -creating webdriver instance-
    driver = webdriver.Chrome(PATH)

    # opening link in webbrowser
    driver.get(link_to_go)
    # wait till the link get loaded
    WebDriverWait(driver, 60)

    # Open Meeting in Window app
    send_keys("{LEFT}")
    send_keys("{ENTER}")

    # -----------------------------------------------------------------------------------------------------------
    # Workaround is needed to open meeting in browser if app is not installed or dont want to open in window app
    # -----------------------------------------------------------------------------------------------------------

    # -----------handelling warnings if any -------------
    warnings.simplefilter('ignore', category=UserWarning)

    # --------Connect to cisco webex meetng app----------
    app = Application().connect(
        title_re=".*Meetings", class_name="wcl_manager1")

    app_window = app.window(title_re=".*Meetings",
                            class_name="wcl_manager1")

    # Close chrome tab and connect to meeting once app is connected
    if app_window.exists():
        driver.close()
        app_window.set_focus()

        time.sleep(10)
        send_keys("{ENTER}")


# Scheduling outlook calender events for 15 minutes
schedule.enter(900, 1, get_calender, ())

while(1):

    schedule.run()
    cal = get_calender()
    meet = list(cal['Start'])
    nowTime = datetime.datetime.now().strftime("%H:%M")
    if nowTime in meet:
        join(cal, nowTime)

#!/usr/bin/env python3

# Chase Theodos
# 02/16/22
# Boss Auto Registration

# External libraries needed - pywin32, requests

import re, os, sys, requests, win32com.client
from tkinter import messagebox, Tk
from datetime import datetime
from time import sleep

# Student ID
username = ''
# Boss password
password = ''
# Registration date format - 2 digit month/day/year
date = '00/00/00'
# 24-hour format - Hour:Min
# ex: 2pm = '14:00'
time = '00:00'
# Quarter - single digit - Fall=1, Winter=2, Spring=3, Summer=4
quarter = '1'
# Class call numbers here
callNumbers = ['', '', '', '', '']
# Change to true to turn on **windows** scheduler feature
enableTask = True


boss = 'http://boss.latech.edu'
url = 'https://boss.latech.edu/ia-bin/tsrvweb.cgi'
# Checks if boss is down
def downDetector():
	g = requests.get(boss)
	if ("UNAVAILABLE" in g.text):
		print('*** Boss is Currently Down ***')
		print('*** Trying Again in One Minute ***')
		return True
	else:
		return False

# Alert box once class registration is attempted
def alert(title, message, kind='info', hidemain=True):
    if kind not in ('error', 'warning', 'info'):
        raise ValueError('Unsupported alert kind.')
    show_method = getattr(messagebox, 'show{}'.format(kind))
    show_method(title, message)

# Loop until registration day/time
def checkTime():
	# Windows scheduler/date check
	if enableTask:
		if date >= datetime.now().strftime("%x"):
			if time > datetime.now().strftime("%H:%M"):
				print('*** Running Task Scheduler ***')
				taskScheduler()

	# Time check
	while True:
		now = datetime.now()
		day = now.strftime("%x")
		current_time = now.strftime("%H:%M")

		if (day < date) or (day == date and current_time < time):
			print('*** Date/time not reached')
			print('*** Sleeping for 1 minute ***')
			sleep(60)
		else:
			main()
			break

def taskScheduler():
	# Scheduler connection
	task = win32com.client.Dispatch('Schedule.Service')
	task.Connect()
	root = task.GetFolder('\\')
	newTask = task.NewTask(0)
	
	# if task exists, delete old, create new
	try:
		root.GetTask('Boss Auto Registration')
		root.DeleteTask('Boss Auto Registration', 0)
		print('*** Deleting Old Scheduler Task ***')
	
	# create new task if none found
	except:
		createTask(root, newTask)
	
	# create new task after deleting previous
	createTask(root, newTask)

def createTask(root, newTask):
	print('*** Creating New Scheduler Task ***')
	# Trigger
	# year, month, day, hour, min
	time2 = [f'20{date[-2:]}',f'{date[1]}',f'{date[3:5]}',f'{time[:2]}',f'{int(time[-2:])}']
	
	# scheduler month formatting
	if date[0] == '0':
		time2[1] = date[1]

	# convert back to ints
	for num in range(len(time2)):
		time2[num] = int(time2[num])

	setTime = datetime(time2[0],time2[1],time2[2],time2[3],time2[4],0,000000)
	trigger_time = 1
	trigger = newTask.Triggers.Create(trigger_time)
	trigger.StartBoundary = setTime.isoformat()

	# Action
	actionExec = 0
	action = newTask.Actions.Create(actionExec)
	action.ID = 'nada'
	# Find python exec
	action.Path = rf'"{sys.exec_prefix}\python.exe"'
	# Find current working directory of boss.py
	action.Arguments = rf'"{os.getcwd()}\{__file__}"'

	# Parameters
	newTask.RegistrationInfo.Description = 'Boss Auto Registration'
	newTask.Settings.Enabled = True
	newTask.Settings.StopIfGoingOnBatteries = False

	# Saving
	createUpdate = 6
	noLogon = 0
	root.RegisterTaskDefinition('Boss Auto Registration', newTask, createUpdate, '', '', noLogon)
	print('*** Task Created Successfully ***')
	sleep(3)
	exit()

def main():
	with requests.session() as s:
		if not downDetector():
			# Login request data
			data = {
				'tserve_host_code': 'HostZero',
				'tserve_tiphost_code': 'TipZero',
				'ConfigName': 'admnmenu',
				# may need to look at this
				'Term': f'20{date[-2:]}{quarter}',
				'tserve_tip_write':'%7C%7CWID%7CSID%7CPIN%7CTerm%7CConfigName',
				'tserve_trans_config': 'astulog.cfg',
				'SID': username,
				'PIN': password,
				'LoginCD': '10'
			}

			# Establish session with credentials
			q = s.post(url, data=data)

			# Check if credentials correct
			if "invalid" in q.text:
				Tk().withdraw()
				alert("Boss.py", "*** Incorrect Username or Password ***\n")
				exit()
			
			# Error tracking
			notAdded = ""

			# POST req to add classes
			for num in callNumbers:
				# Action R = Register
				# Keep url encoding
				addDropData = 'tserve_tip_read_destroy=&tserve_host_code=HostZero' \
					'&tserve_tiphost_code=TipZero&tserve_trans_config=rstureg.cfg' \
					'&tserve_tip_write=%7C%7CWID%7CSID%7CPIN%7CTerm%7CAwdYear%7CAdTyCode%7CSubject%7CCourseID%7CConfigName' \
					'&Callnum=31480&Action=N' \
					'&Callnum=31633&Action=N' \
					'&GrdType=&Credit=' \
					f'&Action=R&Callnum={num}' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum=' \
					'&Action=R&Callnum='
			
				r = s.post(url, data=addDropData)

				courses = r.text

				if "registrations did not occur:" in r.text:
					notAdded+= "\n" + str(num)
			
			# Regex pulls currently registered classes
			classes = re.findall('\"[A-Z]+ [A-Z]+\"|\"[A-Z]+ [A-Z]+ [A-Z]+\"|\"[A-Z]+ [A-Z]+ [A-Z]+ [A-Z]+\"', r.text)
			
			# Removes unwanted regex match
			if len(classes) > 0:
				classes.pop(0)
			else:
				print('*** You are not currently registered for any classes ***')

			# Display class info
			classList = ""
			for course in classes:
				classList += ("\n" + course.replace('\"', '').title())

			# Display courses not added
			if len(notAdded) > 0 :
				Tk().withdraw()
				alert("Boss.py", "*** Unable to add the following: ***\n" + notAdded)

			# Show currently enrolled classes
			Tk().withdraw()
			alert("Boss.py", "*** Current Classes: ***\n" + classList)
			exit()

		# Boss is down
		else:
			sleep(60)
			main()

checkTime()
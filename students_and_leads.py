#Ученики со статусом занимается
#Все остальные плюс все лиды

#Выгрузка учеников со статусом "Занимается"
# -*- coding: utf8 -*-

import requests
import urllib.parse
import datetime
from pathlib import Path
import xlsxwriter
import re                     
import xlrd

both_key = '' #key

getStudents = "https://coddy.t8s.ru/Api/V2/GetStudents"
getEdUnitStudents = "https://coddy.t8s.ru/Api/V2/GetEdUnitStudents"
getEdUnits = "https://coddy.t8s.ru/Api/V2/GetEdUnits"
getLeads = "https://coddy.t8s.ru/Api/V2/GetLeads"

key = '' #key

email = list()

#Leads 

params2 = {'authkey': both_key, 'attached': "false"}
params2 = urllib.parse.urlencode(params2)
print('getStudents', params2)
r2 = requests.get(getLeads, params=params2)
leads = r2.json()['Leads']

for i in range(len(leads)):
	if 'Agents' in list(leads[i].keys()):
		for j in range(len(leads[i]['Agents'])):
			if 'EMail' in list(leads[i]['Agents'][j].keys()):
				if leads[i]['Agents'][j]['UseEMailBySystem'] == True:
					email.append(leads[i]['Agents'][j]['EMail'])
					# print(leads[i])
					# print()

q1 = len(email)

#Students
params2 = {'authkey': key}
params2 = urllib.parse.urlencode(params2)
print('getStudents', params2)
r2 = requests.get(getStudents, params=params2)
students = r2.json()['Students']

for i in range(len(students)):
	if 'Status' in list(students[i].keys()):
		if students[i]['Status'] != 'Занимается':
			if 'Agents' in list(students[i].keys()):
				for j in range(len(students[i]['Agents'])):
					if 'EMail' in list(students[i]['Agents'][j].keys()):
						if students[i]['Agents'][j]['UseEMailBySystem'] == True:
							email.append(students[i]['Agents'][j]['EMail'])
							# print(students[i])
							# print()
	else:
		if 'Agents' in list(students[i].keys()):
			for j in range(len(students[i]['Agents'])):
				if 'EMail' in list(students[i]['Agents'][j].keys()):
					if students[i]['Agents'][j]['UseEMailBySystem'] == True:
						email.append(students[i]['Agents'][j]['EMail'])
						# print(students[i])
						# print()

q2 = len(email) - q1

print(q1)
print(q2)
print(len(leads))
print(len(students))

lm = len(email) // 3000
lmo = len(email) % 3000

# for i in range(1, lm+1):
# 	workbook = xlsxwriter.Workbook(f'students_and_leads-{i}.xlsx')
# 	worksheet = workbook.add_worksheet()
# 	bold = workbook.add_format({'bold': True})
# 	temp = i * 3000
# 	if i == 1:
# 		for j in range(temp):
# 			worksheet.write(f'A{j+2}', email[i])
# 	else:
# 		for j in range(temp1, temp):
# 			worksheet.write(f'A{j+2}', email[i])
# 	temp1 = temp
# 	workbook.close()

# workbook = xlsxwriter.Workbook(f'students_and_leads-last.xlsx')
# worksheet = workbook.add_worksheet()
# bold = workbook.add_format({'bold': True})
# for i in range(temp1, len(email)):
# 	worksheet.write(f'A{i+2}', email[i])

workbook = xlsxwriter.Workbook(f'students_and_leads.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
for i in range(len(email)):
	worksheet.write(f'A{i+2}', email[i])

workbook.close()
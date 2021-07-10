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

key = '' #key

params2 = {'authkey': key, 'statuses': 'Занимается'}
params2 = urllib.parse.urlencode(params2)
# print('getStudents', params2)
r2 = requests.get(getStudents, params=params2)
students = r2.json()['Students']

email = list()

for i in range(len(students)):
	if students[i]['Status'] == 'Занимается':
		if 'Agents' in list(students[i].keys()):
			for j in range(len(students[i]['Agents'])):
				if 'EMail' in list(students[i]['Agents'][j].keys()):
					if students[i]['Agents'][j]['UseEMailBySystem'] == True:
						email.append(students[i]['Agents'][j]['EMail'])
						print(students[i])
						print()

workbook = xlsxwriter.Workbook(f'all_students.xlsx')
worksheet = workbook.add_worksheet()

bold = workbook.add_format({'bold': True})


for i in range(len(email)):
	worksheet.write(f'A{i+2}', email[i])

workbook.close()
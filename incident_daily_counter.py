import resilient
import time
import json
import re
import csv
import os
import smtplib
import ssl
import xlwt
import xlsxwriter
import email
import email.header
import email.mime.multipart
import datetime as dt
import locale
import calendar

from calendar import monthrange
from datetime import datetime, timedelta
from xml.dom import minidom
from xlwt import Workbook

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def main():

	yesterday   = dt.date.today() - timedelta(1)
	today	   = dt.date.today()

	start_date  = int(time.mktime(yesterday.timetuple())*1000)
	end_date	= int(time.mktime(today.timetuple())*1000)

	list_incident = export_incident_from_resilient()
	
	info = {}
	info['date'] = dt.date.strftime(yesterday, '%Y%m%d')
	info['type'] = 'Daily'

	create_xlsx(list_incident, 'Daily Report - Incidents {}.xlsx'.format(info['date']))

	day = dt.date.strftime(dt.date.today(), '%A')

def get_week_info(date):
	get_first_date = date.replace(day=1)
	date_range = calendar.monthrange(date.year, date.month)
	date_and_day = [None] * (date_range[1] + 1)
	week = []
	list_week = {}
	counter = 1
	last_three_week = ['Friday', 'Saturday', 'Sunday']
	
	for i in range(1, date_range[1] + 1):
		date_and_day[i] = date.replace(day=i).strftime('%A')

		week.append(i)
		
		if i == 1 and date_and_day[i] in last_three_week:
			counter = 0 

		if date_and_day[i] in 'Sunday' or i == date_range[1]:
			list_week['W' + str(counter)] = week
			week = []
			counter = counter + 1

	return list_week 

def get_week_number(date):
	list_week = get_week_info(date)

	week_number = 'W'

	for key in list_week:
		if key == 'W0' and date.day in list_week['W0']:
			new_list_week = get_week_info(date.replace(day=1) - timedelta(days=1))
			temp = []
			for n in new_list_week:
				temp.append(n) 

			week_number = temp[len(temp) - 1]
			continue

		if date.day in list_week[key]:
			if len(list_week[key]) < 4:
				key = 'W1'

			week_number = key

	return week_number

def export_incident_from_resilient():

	TAG_RE = re.compile(r'<[^>]+>')
	parser = resilient.ArgumentParser(config_file=resilient.get_config_file())
	opts = parser.parse_args()

	client = resilient.get_client(opts)

	# prepare the future
	body = {
		"filters": [
			{
				"conditions": [
					{
						"field_name": "plan_status",
						"method": "equals",
						"value": "Active"
					}
				],
				"logic_type" : "any"
			}
		],
		'start': 0,
		'length': 1000,
		'sorts': [
			{
				'field_name': 'id',
				'type': 'asc'
			}
		]
	}

	uri = "/incidents/query_paged?return_level=normal"
	incidents = client.post(uri, body)
	incident_ids = []

	while incidents.get('data'):
		data = incidents.get('data')
		
		for result in data:
			incident_ids.append(result['id'])
			
		body['start'] = len(data) + body['start']

		incidents = client.post(uri, body)

	incident_ids.sort()

	uri = "/types"

	types = client.get(uri)

	severity_code = {}
	incident_type_ids = {}
	assigned_group = {}
	plan_status = {}
	list_incident = []

	for inc_id in incident_ids:
		list_incident_ahey = []

		uri = "/incidents/{}".format(inc_id)
		the_incident = client.get(uri)

		yes = datetime.strptime(datetime.fromtimestamp(int(the_incident['create_date'])/1000).strftime('%Y-%m-%d %H:%M:%S'), '%Y-%m-%d %H:%M:%S')
		days = yes-datetime.now()
		days = -1*days.days
		uri = '/incidents/{}'.format(inc_id)

		incident = client.get(uri)

		# Record
		incident['properties']['days_alive'] = days

		client.put(uri, incident)

	return list_incident

def create_xlsx(data, filename):
	output_path		= os.path.realpath('incident_log')

	workbook 			= xlsxwriter.Workbook(os.path.join(output_path, filename))
	incidents_sheet 	= workbook.add_worksheet('Incidents')
	total_sheet			= workbook.add_worksheet('Total')

	header_text		= workbook.add_format({
		'align': 'center',
		'valign': 'vcenter',
		'font_name': 'Calibri',
		'font_size': 11,
		'border': 1,
		'bold': 1,
		'fg_color': '#7BC0FF'})

	text			= workbook.add_format({
		'font_name': 'Calibri',
		'font_size': 11,
		'num_format': '#,##0',
		'border': 1})	

	format_text = {'text': text, 'header': header_text}

	incidents_sheet.set_column(0, 0, 16)  
	incidents_sheet.set_column(1, 1, 10)  
	incidents_sheet.set_column(2, 2, 50)  
	incidents_sheet.set_column(3, 3, 14)  
	incidents_sheet.set_column(4, 4, 14)
	incidents_sheet.set_column(5, 5, 16)
	incidents_sheet.set_column(6, 6, 7)
	incidents_sheet.set_column(7, 7, 18)
	incidents_sheet.set_column(8, 8, 18)
	incidents_sheet.set_column(9, 9, 18)
	incidents_sheet.set_column(10, 10, 6)
	incidents_sheet.set_column(11, 11, 66)

	total_rows = write_data_in_sheet(incidents_sheet, data, format_text)

	total_sheet.set_column(0, 0, 12)  
	total_sheet.set_column(1, 1, 8)  

	total_sheet.write('A1', 'Status', header_text)
	total_sheet.write('B1', 'Count', header_text)

	total_sheet.write('A2', 'Open', text)
	total_sheet.write('A3', 'Close', text)

	total_sheet.write('B2', '=COUNTIF(Incidents!H2:H{},"Active")'.format(total_rows), text)
	total_sheet.write('B3', '=COUNTIF(Incidents!H2:H{},"Closed")'.format(total_rows), text)

	total_sheet.write('A4', 'Total', text)
	total_sheet.write('B4', '=SUM(B2:B3)', text)

	workbook.close()

def write_data_in_sheet(worksheet, data, format_text):

	row_counter	= 0
	col_counter = 0
	
	formatting = format_text['header']

	for row in data:
		col_counter = 0
		for cell in row:
			
			if row_counter > 0:
				formatting = format_text['text']

			worksheet.write(row_counter, col_counter, cell, formatting)
			col_counter += 1
		
		row_counter += 1

	return row_counter

if __name__ == "__main__":
	main()
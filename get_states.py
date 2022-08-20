#!/usr/bin/env python

import json
import sys
import requests
import pandas
from openpyxl import load_workbook
import xlsxwriter

# get list of states with lattitude and longitude for request
states_df = pandas.read_csv("state_locations.tsv", sep="\t")


# open excel to write results
filename_excel = "intersocietal_nuclear_pet_locations_by_state.xlsx"
writer = pandas.ExcelWriter(filename_excel, engine='xlsxwriter')

# get data for each state
for i, state_row in states_df.iterrows():

	state = dict(state_row)
	print(f"getting data for state {i}:", state)
	
	# fetch data from website 
	# (I got this URL from inspecting the form response at https://intersocietal.org/iac-accredited-facility-locator/)
	url = f"https://www.intersocietal.org/proxyServices/MainService.aspx/GetLocations?callback=mapload&origLat={state['LATT']}&origLng={state['LONG']}&origAddress={state['ABBR']}&_=1660948634742"
	print(f"fetching data from url: {url}")
	response = requests.get(url)

	# process response
	# strip of wrapper function text 'mapload(...)';
	print(f"response", response.status_code)
	state_json = response.text.replace("mapload(", "").replace(");", "")
	
	# save JSON data from response
	filename_json = state['ABBR'] + ".json"
	with open(filename_json, mode="w") as file_json:
		print(f"writing results json data to file: {filename_json}")
		file_json.writelines(state_json)

	data = json.loads(state_json)
	
	# load JSON to pandas dataframe
	print("load dataframe")
	state_df = pandas.json_normalize(data=data) 

	# dump raw data to CSV file
	filename_csv = state['ABBR'] + ".csv"
	print(f"writing results data to csv: {filename_csv}")
	state_df.to_csv(filename_csv)

	# abbreviated data for Excel workboook
	short_df = pandas.DataFrame(data=state_df, columns=["name", "address", "address2", "city", "state", "postal", "country", "phone", "web"])

	print(short_df)

	
	# write to excel workbook
	worksheet_name = state['ABBR']
	print(f"writing worksheet {worksheet_name} to excel: {filename_excel}")

	workbook = xlsxwriter.Workbook(filename_excel)
	workbook.add_worksheet(worksheet_name)

	short_df.to_excel(writer, sheet_name=worksheet_name, index=False)

writer.save()
print("done")

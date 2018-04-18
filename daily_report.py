import requests, json, openpyxl
import datetime as dt
import user

def create_workbook():
	wb = openpyxl.Workbook()
	ws = wb.active
	ws.title = "Epic"

	ws["A1"].value = "Key"
	ws["B1"].value = "Summary"
	ws["C1"].value = "Title"
	ws["D1"].value = "Version"
	ws["E1"].value = "Story Points"
	ws["F1"].value = "NP SP"
	ws["G1"].value = "IP SP"
	ws["H1"].value = "Fix SP"

	return wb

def save_workbook(wb):
	date = "{year}{month}{day}".format(year=dt.datetime.now().year,
									month=dt.datetime.now().month if dt.datetime.now().month >= 10 else "0"+str(dt.datetime.now().month),
									day = dt.datetime.now().day if dt.datetime.now().day >= 10 else "0"+str(dt.datetime.now().day))
	wb.save("SENT2 Status {}.xlsx".format(date))

def edit_workbook(wb, dic):  #, epic):
	ws = wb.active

	row = 2
	for i in range(len(dic['issues'])):
		ws["A{}".format(row)].value = dic['issues'][i]['key']
		ws["B{}".format(row)].value = dic['issues'][i]['fields']['summary']
		ws["C{}".format(row)].value = dic['issues'][i]['key'] + ' ' + dic['issues'][i]['fields']['summary']
		if dic['issues'][i]['key'] in ['SENT2-23', 'SENT2-22', 'SENT2-15', 'SENT2-14', 'SENT2-12', 'SENT2-11']:
			ws["D{}".format(row)].value = "SENT 1.0"
		elif dic['issues'][i]['key'] in ['SENT2-9', 'SENT2-8', 'SENT2-7', 'SENT2-6', 'SENT2-5', 'SENT2-4', 'SENT2-3', 'SENT2-2', 'SENT2-10', 'SENT2-1']:
			ws["D{}".format(row)].value = "SENT 2.0"
		else:
			ws["D{}".format(row)].value = ""
		ws["E{}".format(row)].value = dic['issues'][i]['fields'].get('customfield_11213','')


		# Child Stuff
		# res = get_child_info(dic['issues'][i]['key'])
		# ws["F{}".format(row)].value = res['np']
		# ws["G{}".format(row)].value = res['ip']
		# ws["H{}".format(row)].value = res['cp']

		row += 1


	ws["C{}".format(row)].value = "Security Roll up"
	ws["F{}".format(row)].value = '=SUMIF(B2:B{row},"Security*",F2:F{row})'.format(row=row)
	ws["G{}".format(row)].value = '=SUMIF(B2:B{row},"Security*",G2:G{row})'.format(row=row)
	ws["H{}".format(row)].value = '=SUMIF(B2:B{row},"Security*",H2:H{row})'.format(row=row)
	ws["D{}".format(row)].value = "SENT 1.0"


	tab = openpyxl.worksheet.table.Table(displayName="Epic",ref="A1:H{}".format(row))
	ws.add_table(tab)



def get_child_info(key):
	res = {'np': 0, 'ip': 0, 'cp': 0} # Not Started, In Progress, Completed
	info = json.loads(requests.get("https://{jira}/rest/api/2/search?jql=cf[13300]in({key})&maxResults=300&fields=customfield_11213,status,resolution".format(key=key,jira=user.company_jira), auth=(user.username, user.password)).content)

	for i in range(len(info['issues'])):
		sp = info['issues'][i]['fields']['customfield_11213']

		# Check if resolution = Fixed, then update res['cp']
		if info['issues'][i]['fields']['resolution'] != None:
			if info['issues'][i]['fields']['resolution']['name'] == 'Fixed':
				if type(sp) == float:
					res['cp'] += sp
		# Else, check if status = In Progress, then update res['ip']
		elif info['issues'][i]['fields']['status']['name'] == 'In Progress':
			if type(sp) == float:
				res['ip'] += sp #X

		# Else, it's not worked on, so update res['np']
		else:
			if type(sp) == float:
				res['np'] += sp

	return res
def get_filter(filter):
	r = json.loads(requests.get("https://{jira}/rest/api/2/filter/{filter}".format(filter=filter,jira=user.company_jira),auth=(user.username,user.password)).content)
	return r['jql']

def get_epics(jql_raw):
	open_paren, close_paren = jql_raw.find('('), jql_raw.find(')')
	jql_raw = jql_raw[open_paren+1:close_paren]
	epics = jql_raw.replace('%2C',',')
	new_jql = "https://{jira}/rest/api/2/search?jql=key+in+({epics})".format(epics=epics,jira=user.company_jira)
	new_jql += "&fields=summary,customfield_11213,customfield=11701&maxResults=200"

	jira_dic = json.loads(requests.get(new_jql, auth=(user.username,user.password)).content)

	return jira_dic

def collect_stories(wb):
	ws2 = wb.create_sheet(title="Story")

	ws2["A1"].value = "Key"
	ws2["B1"].value = "Summary"
	ws2["C1"].value = "Story Points"	#customfield_11213
	ws2["D1"].value = "Sprint"			#customfield_11701
	ws2["E1"].value = "Team/s"			#customfield_14400
	ws2["F1"].value = "Epic Link"		#customfield_13300
	ws2["G1"].value = "Status"			


	url = "https://{jira}/rest/api/2/search?jql={filter}&maxResults=5000&fields=summary,customfield_11213,customfield_11701,customfield_14400,customfield_13300,status".format(filter=get_filter(106654),jira=user.company_jira)
	dic = json.loads(requests.get(url,auth=(user.username,user.password)).content)

	row = 2
	for i in range(len(dic['issues'])):
		ws2["A{}".format(row)].value = dic['issues'][i]['key']
		ws2["B{}".format(row)].value = dic['issues'][i]['fields']['summary']
		# ws2["C{}".format(row)].value = dic['issues'][i]['key'] + ' ' + dic['issues'][i]['fields']['summary']
		ws2["C{}".format(row)].value = dic['issues'][i]['fields'].get('customfield_11213','')

		if dic['issues'][i]['fields'].get('customfield_11701',''):
			sprint = dic['issues'][i]['fields']['customfield_11701']
			st = sprint[0].find('name=')
			end = sprint[0][st:].find(',')

			sprint = sprint[0][st+5:st+end]
			ws2["D{}".format(row)].value = sprint
		# ws2["D{}".format(row)].value = dic['issues'][i]['fields'].get('customfield_11701','')
		
		team = dic['issues'][i]['fields'].get('customfield_14400','')
		if team != None:
			ws2["E{}".format(row)].value = team[0]
		# ws2["E{}".format(row)].value = dic['issues'][i]['fields'].get('customfield_14400[0]','')
		ws2["F{}".format(row)].value = dic['issues'][i]['fields'].get('customfield_13300','')
		ws2["G{}".format(row)].value = dic['issues'][i]['fields']['status']['name']
		row += 1


if __name__ == "__main__":
	wb = create_workbook()
	fil = get_filter(106654)
	epics = get_epics(fil)
	edit_workbook(wb, epics)
	# collect_stories(wb)
	save_workbook(wb)	

"""
Add in:
	- Status on the epics
	- Have another worksheet for all the current sprint information.

"""
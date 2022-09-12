import requests
import base64
import argparse
from datetime import date 
import csv
import pandas as pd
import getpass
#import win32com.client as win32
from tqdm import tqdm
import time

def fetch_scan_id():

	url = "https://qualysapi.qg2.apps.qualys.com/api/2.0/fo/scan/?action=list"
	payload={
	'launched_after_datetime': date.today(),
	'action': 'list',
	'type': 'Scheduled',
	#'state': 'Running',
	}
	files=[
	
	]
	headers = {
	  'X-Requested-With': 'QualysPostman',
	  'Authorization': auth
	}
	
	response = requests.request("POST", url, headers=headers, data=payload, files=files)
	print(date.today())
	print(response.status_code)
	if(response.status_code != 200):
		print(response.text)
		quit()
	s = response.text
	
	titletag = "<TITLE>"
	titleendtag = "</TITLE>"
	start = '<REF>'
	end = '</REF>'

	
	for item in s.split(end):
		if start in item:
			#print(item [item.find(start)+len(start) : ])
			scan_id.append(item [item.find(start)+len(start) : ])
	
	
	for item in s.split(titleendtag):
		if titletag in item:
			#print(item [item.find(titletag)+len(titletag) : ])
			scan_name.append(item [item.find(titletag)+len(titletag) : ])
	

	for i in range(0,len(scan_name)):
		
		if(scan_name[i] == "<![CDATA[[Daily]: External Scan]]>" or scan_name[i] == "<![CDATA[LMK External In Scope - Daily Scan [All Cloud]]]>"):
			print("processing scan with ID", scan_id[i],"with name", scan_name[i])
			fetch_scan_report(scan_name[i], scan_id[i])


def fetch_scan_report(scan_name, scan_id):

	url = "https://qualysapi.qg2.apps.qualys.com/api/2.0/fo/scan/?action=fetch"
	i=0


	payload={
	'scan_ref': scan_id,
	'output_format': 'csv_extended'
	}

	files=[
	
	]
	headers = {
	  'X-Requested-With': 'QualysPostman',
	  'Authorization': auth
	}
	
	response = requests.request("POST", url, headers=headers, data=payload, files=files)
	print(date.today())
	print("Received response, processing scan with ID: "+ scan_id +" with name: "+scan_name)
	print(response.status_code)
	if(response.status_code != 200):
		print(response.text)
	#if (response.text).find("unable"):
	#	print("come back in evening")
	#	print(response.text)
	#if ((response.text).find("<?xml")):
	#	print("Scan results not available as scans are still running. exiting......")
	#	quit()
	#else:
	if(scan_name == "<![CDATA[[Daily]: External Scan]]>"):
		export_csv(response.text,(d.replace("-","_")+"_Daily_Scan_HAL"))
	else:
		export_csv(response.text,(d.replace("-","_")+"_Daily_Scan_LMK"))

	#print(response.text)


def export_csv(data,name):
	file = open("scan_HAL.csv","w")
	file.write(data)
	file.close()
	print("data written successfully in initial file, now filtering data......")
	with open("scan_HAL.csv", newline='') as in_file:
		with open("first_edit.csv", 'w', newline='') as out_file:
			writer = csv.writer(out_file)
			for row in csv.reader(in_file):
				if row:
					writer.writerow(row)
	df = pd.read_csv('first_edit.csv',encoding='cp1252',skiprows = 5)
	df =  df[(df.Severity >= 4) & (df.Type == "Vuln")]
	filen = name+".csv"
	df.to_csv(filen)
	print('filtered data in '+name+'.csv.....')
	#print(df)
	if(filen.find("HAL") != -1):
		print("Launching report for HAL selected IPs......")

		df = pd.read_csv(filen)
		#df = pd.DataFrame(file)
		ips = ",".join(list(df.IP.unique()))
		print(ips)
		fetch_report(filen,ips)

	#print("sev 4 & 5 are:")
	#reader = csv.reader(open(r"scan_HAL.csv"),delimiter=',')
	#filtered = filter(lambda p: '45' == p[8], reader)
	#print(list(filtered))

'''
def send_email(filename,name):
	s = "xxxxxxxxx"
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.To = 'xxxxxxxxx'
	mail.Subject = 'DO NOT REPLY -- Daily External Scan '+name
	mail.Body = """
	Hi All,

	PFA Daily scan report for External facing assets/servers having sev 4 & 5 vulnerabilites.
	Kindly do the needfull.

	Please note that this is a auto-generated mail.

	Thanks,
	Kshitiz Thakur
	""" 
	#mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

	# To attach a file to the email (optional):
	attachment  = s+filename
	mail.Attachments.Add(attachment)
	mail.Send()
	print('mail sent successfully')
	'''


def fetch_report(filename,ips):
	url = "https://qualysapi.qg2.apps.qualys.com/api/2.0/fo/report/?action=launch"
	payload={
	'report_type': 'Scan',
	'ips': ips,
	'template_id': 2365844,
	'report_title': 'Daily_External_Scan_Report_Selected_IPs',
	'output_format': 'csv'
	}
	files=[
	
	]
	headers = {
	  'X-Requested-With': 'QualysPostman',
	  'Authorization': auth
	}
	
	response = requests.request("POST", url, headers=headers, data=payload, files=files)
	print("Launching HAL Report.......")
	print(response.status_code)
	s = response.text
	start = "<TEXT>"
	end = "</TEXT>"
	vstart = "<VALUE>"
	vend = "</VALUE>"

	for item in s.split(end):
	  if start in item:
	    #print(item [item.find(start)+len(start) : ])
	    print(item [item.find(start)+len(start) : ])

	for item in s.split(vend):
	  if start in item:
	    #print(item [item.find(start)+len(start) : ])
	    r_id = (item [item.find(vstart)+len(vstart) : ])

	for i in tqdm (range (101),
			desc="Processing Report...",
			ascii=False, ncols=75):
	time.sleep(0.3)


	url = "https://qualysapi.qg2.apps.qualys.com/api/2.0/fo/report/?action=fetch"
	payload={
	'id' = r_id
	}
	files=[
	
	]
	headers = {
	  'X-Requested-With': 'QualysPostman',
	  'Authorization': auth
	}
	
	response = requests.request("POST", url, headers=headers, data=payload, files=files)

	print("Writing Data to File.......")
	file = 	open("REPORT_"+filename,"w")
	file.write(s)
	file.close()
	print("Success!")



if __name__ == '__main__':
	
	parser = argparse.ArgumentParser()
	parser.add_argument('-uname', type=str, required=True)
	#parser.add_argument('-passwd', type=str, required=True)
	x = getpass.getpass("Enter Password:")
	args = parser.parse_args()
	#<![CDATA[[Daily]: External Scan]]>	
	#<![CDATA[LMK External In Scope - Daily Scan [All Cloud]]]>
	UP = args.uname+":"+x
	hashv = base64.b64encode(bytes(UP,'utf-8'))
	auth = 'Basic '+hashv.decode('ascii')
	scan_id = []
	scan_name = []
	d = str(date.today())

	fetch_scan_id()

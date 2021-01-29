#PreRequisite:
#Grab receipts under "Business" category are sent to outlook
#Python 3
#pip install comtypes.client
#pip install PyPDF2

import win32com.client #Used to connect to Outlook
import os #For traversing directories
import comtypes.client #
import PyPDF2
from datetime import datetime, date
import csv

def monthToOrdinal(month):
	monthMap={
	'January':'1', 
	'February':'2', 
	'March':'3', 
	'April':'4', 
	'May':'5', 
	'June':'6', 
	'July':'7', 
	'August':'8', 
	'September':'9', 
	'October':'10', 
	'November':'11', 
	'December':'12'
	}
	return monthMap[month]

def getEmails(year,month,day):
	#Ensure the path to place the emails is created, if not create the path
	try:
		os.chdir("C:\GrabReceipts")
	except:
		os.mkdir("C:\GrabReceipts")
		os.chdir("C:\GrabReceipts")

	#retrieve the emails
	messages = win32com.client.Dispatch('outlook.application').GetNameSpace("MAPI").GetDefaultFolder(6).Items

	#filter the messages to be only from Grab and after the stipulated date
	received_dt = date(year,month,day)
	received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
	messages = messages.Restrict("[Subject] = '[Business] Your Grab E-Receipt'")
	messages = messages.Restrict("[ReceivedTime] >= '" +received_dt+"'")

	with open('claims.csv',mode='w') as claim_file:
		claim_writer = csv.writer(claim_file,delimiter=',')
		for message in messages:
			#save the emails in word format
			msgname= message.Subject + "_"+ message.ReceivedTime.strftime('%d-%m-%Y - %H-%M') + ".doc"
			message.SaveAs(os.getcwd()+'//'+msgname,4)

			#convert the word to PDF format using comtype module
			word = comtypes.client.CreateObject('Word.application')
			doc = word.Documents.Open(os.getcwd()+'//'+msgname)
			doc.SaveAs(os.getcwd()+'//'+msgname +".pdf", FileFormat=17)
			doc.Close()
			word.Quit()

			#Delete the word doc
			os.remove(os.getcwd()+'//'+msgname)

			#Extract the content of the PDF
			pdfFileObj = open(os.getcwd()+'//'+msgname +".pdf", 'rb')
			pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
			pageObj = pdfReader.getPage(0)
			extractedList = pageObj.extractText().split('\n')
			price = extractedList[59]
			dayOfRec, monthOfRec, yearOfRec = extractedList[8].split(" ")[0:3]
			monthOfRec = monthToOrdinal(monthOfRec)

			#Write the content onto thereceipt
			claim_writer.writerow([dayOfRec + '/' + monthOfRec + '/' + yearOfRec,price])


if __name__=="__main__":
	print("Enter the date from which you wish to require the receipt thereof from. For Example, if '2021/1/1' is entered, all the receipts beginning from 1 jan 2021 would be retrieved.")
	year,month,day = list(map(int,input("Please enter the date of the earliest receipts(YYYY/MM/DD):").split("/")))
	getEmails(year,month,day)
	# getEmails(2020,12,29)

# PreRequisite:
# Grab receipts under "Business" category are sent to outlook
# Python 3
# pip install comtypes.client
# pip install PyPDF2


import win32com.client  # Used to connect to Outlook
import os  # For traversing directories
import comtypes.client  #
import PyPDF2
from datetime import datetime, date
import csv


def isDigit(x):
    try:
        float(x)
        return True
    except ValueError:
        return False

#Look for the string Total Paid and then find the next floating value after that
def lookForTotalPaid(extractedList):
    i = 0
    while extractedList[i] != 'Total Paid':
        i += 1

    while isDigit(extractedList[i]) is False:
        i += 1
    return extractedList[i]


def monthToOrdinal(month):
    monthMap = {
        'January': '1',
        'February': '2',
        'March': '3',
        'April': '4',
        'May': '5',
        'June': '6',
        'July': '7',
        'August': '8',
        'September': '9',
        'October': '10',
        'November': '11',
        'December': '12'
    }
    return monthMap[month]


def getEmails(year, month, day):
    # Ensure the path to place the emails is created, if not create the path
    try:
        os.chdir("C:\GrabReceipts")
    except:
        os.mkdir("C:\GrabReceipts")
        os.chdir("C:\GrabReceipts")

    # retrieve the emails
    messages = win32com.client.Dispatch('outlook.application').GetNameSpace(
        "MAPI").GetDefaultFolder(6).Items

    # filter the messages to be only from Grab and after the stipulated date
    received_dt = date(year, month, day)
    received_dt = received_dt.strftime('%d/%m/%Y %H:%M %p')
    print(received_dt)

    messages = messages.Restrict(
        "[Subject] = '[Business] Your Grab E-Receipt'")
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

    with open('claims.csv', mode='w') as claim_file:
        claim_writer = csv.writer(claim_file, delimiter=',')
        for message in messages:
            # save the emails in word format
            msgname = "R " + \
                message.ReceivedTime.strftime('%d-%m-%y %H-%M')
            message.SaveAs(os.getcwd() + '//' + msgname + ".doc", 4)

            # convert the word to PDF format using comtype module
            word = comtypes.client.CreateObject('Word.application')
            doc = word.Documents.Open(os.getcwd() + '//' + msgname + ".doc")
            doc.SaveAs(os.getcwd() + '//' + msgname + ".pdf", FileFormat=17)
            doc.Close()
            word.Quit()

            # Delete the word doc
            os.remove(os.getcwd() + '//' + msgname + ".doc")

            # Extract the content of the PDF
            pdfFileObj = open(os.getcwd() + '//' + msgname + ".pdf", 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            pageObj = pdfReader.getPage(0)
            extractedList = pageObj.extractText().split('\n')
            price = lookForTotalPaid(extractedList)
            dayOfRec, monthOfRec, yearOfRec = extractedList[8].split(" ")[0:3]
            monthOfRec = monthToOrdinal(monthOfRec)

            # Write the content onto thereceipt
            claim_writer.writerow(
                [dayOfRec + '/' + monthOfRec + '/' + yearOfRec, price])

            pdfFileObj.close()
            os.rename(os.getcwd() + '//' + msgname + ".pdf",
                      os.getcwd() + '//' + msgname + " $" + price + ".pdf")


if __name__ == "__main__":
    # print("Enter the date from which you wish to require the receipt thereof from. For Example, if '2021/1/1' is entered
    # all the receipts beginning from 1 jan 2021 would be retrieved.")
    year,month,day = list(map(int,input("Please enter the date of the earliest receipts(YYYY/MM/DD):").split("/")))
    getEmails(year,month,day)
    # getEmails(2021, 1, 12)

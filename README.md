# GrabReceiptsFromOutlook

Background:
A part of the job is to do recurring transportation claims - Grab ride claims. This scripts helps to ease the pain of making the claims by doing some pre processing work

Features:
1. Pulls Grab Receipts from Microsoft outlook and saves the receipts in PDF format on "C:\GrabReceipts" from a given date.
   Example: If 2021/1/1 is given, it will pull all receipts received from 2021/1/1 to the present date.
2. Extract information from the receipts and saves this information in a csv file - "C:\GrabReceipts\claims.csv"
   a. Date of Receipt
   b. Value of Receipt

PreRequisite:
1.Grab receipts under "Business" category are sent to your outlook
2.Python 3
3.Run "pip install comtypes.client" in order for the script to work
4.pip "pip install PyPDF2" in order for the script to work

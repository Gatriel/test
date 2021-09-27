from googleapiclient.discovery import build
from google.oauth2 import service_account
from math import ceil

################################################################################################################
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json' #keys.json is the json file containing the information needed to access the service acc
creds = None #intialize creds (creds stand for credentials)
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
SPREADSHEET_ID = '1vO_XyKJafVU4r4Rr69eaWR1gd-CUzoMMMxaAIssgob4'  #saves the ID of the spreadsheet we want to manipulate
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()  #calls the Sheets API
#################################################################################################################


students = 0  #intializes a variable to count how many students there are, starting at 0. Will also be used as an iterator
r = "engenharia_de_software!C4:F4"  #initializes the range in the first line with useful values
info = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=r).execute()  #initializes the info with the initial range

while info.get('values', []):  #loop that goes until the last registered value in the spreadsheet
    students+=1  #increments iterator so we can change the range variable and also count how many students there are
    r = "C" + str(4+students) + ":F" + str(4+students)  #updates range to the next line of values
    info = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=r).execute()  #updates the info with the new range


result = sheet.values().get(range="A1:F"+str(4+students),spreadsheetId=SPREADSHEET_ID).execute() #gets all the info from the spreadsheet
values = result.get('values', []) #makes a list of lists with the values
classes = values[1][0] #gets the string containing the total number of classes
classes = int(classes[len(classes)-3:len(classes)])  #gets the 3 last characters of the string, where the number of classes is, and turns it into an int


for s in range(3, 3+students):  #goes through all the students in the spreadsheet

    attendance = 1-(float(values[s][2])/float(classes))  #calculates the attendance of each student in the list

    if attendance < 0.75:  #if the student didnt come to more than 25% of the classes, updates his status in the sheet
        assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                   range="G"+str(s+1), valueInputOption="USER_ENTERED",
                   body={"values": [["Reprovado por Falta"]]}).execute()
        assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                       range="H" + str(s + 1), valueInputOption="USER_ENTERED",
                                       body={"values": [["0"]]}).execute()
    else:  #if he has enough attendance to be able to pass, calculates the grade average
        average = (float(values[s][3])+float(values[s][4])+float(values[s][5]))/3  #calculates the average of the grades
        average = ceil(average)  #rounds the average up

        if average < 50:
            assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                           range="G" + str(s + 1), valueInputOption="USER_ENTERED",
                                           body={"values": [["Reprovado por Nota"]]}).execute()
            assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                           range="H" + str(s + 1), valueInputOption="USER_ENTERED",
                                           body={"values": [["0"]]}).execute()
        elif average < 70:
            feg = 100 - average  #calculates the grade needed in the final exam (final exam grade)
                                 #I decided to keep the feg between 0 and 100, not 0 and 10, because
                                 #the other grades are also from 0 to a 100
            assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                           range="G" + str(s + 1), valueInputOption="USER_ENTERED",
                                           body={"values": [["Exame Final"]]}).execute()
            assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                           range="H" + str(s + 1), valueInputOption="USER_ENTERED",
                                           body={"values": [[str(feg)]]}).execute()
        else:
            assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                           range="G" + str(s + 1), valueInputOption="USER_ENTERED",
                                           body={"values": [["Aprovado"]]}).execute()
            assign = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                           range="H" + str(s + 1), valueInputOption="USER_ENTERED",
                                           body={"values": [["0"]]}).execute()

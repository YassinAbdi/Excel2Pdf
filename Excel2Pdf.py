# Import pandas
import pandas as pd
import numpy as np
import sys
from WritePdfFile import CreateApplicantPdf
from WritePdfFile import CreateFolderForPdf
import datetime
import requests
import json

KEY_LENGTH = 16


def Excel2Pdf(spreadsheetName, TemplateFileName1 ,TemplateFileName2):
    foundError = False
    applicants = []
    index = 0
    windowsInvalidCharters=[chr(92),"/",":","*","?","<",">","|"] # chr(92) is \ but python does allow for string

    # Load spreadsheet
    xl = pd.ExcelFile(spreadsheetName)

    # Load a sheet into a DataFrame by name: df1
    df = xl.parse('Sheet1')
    data = df.to_numpy()
    CorrectIndex = np.array(['RECIPIENT NAME (FIRST, MI, LAST)', 'MA Member #Date of Birth',
    'PCA Name(first, MI, Last', 'PCA NPI/UMPI'])
    columns=df.columns
    for number in range(len(CorrectIndex)):
        if(CorrectIndex[number] != columns[number]):
            foundError = True

    if(foundError == True):
        input("ERROR_01:Could not file corrent index please contact: 612-703-9647")

    # create path for folders
    date = datetime.datetime.now()
    path = CreateFolderForPdf(date)

    # grab all of the applicants after the starter index
    for row in data:            
        applicants.append(row)
        index += 1
    
    for applicant in applicants:
        string0 = str(applicant[0])
        string1 = str(applicant[1])
        string2 = str(applicant[2])
        string3 = str(applicant[3])
        #checking if cell isn't empty
        if(string0 == "nan"):
            applicant[0] = ""
            print("empty cell")
        if(string1 == "nan"):
            applicant[1] = ""
            print("empty cell")
        if(string2 == "nan"):
            applicant[2] = ""
            print("empty cell")
        if(string3 == "nan"):
            applicant[3] = ""
            print("empty cell")

        #checking if cell doesn't contain invalid characters
        for item in windowsInvalidCharters:
            if(string0.find(item) == True):
                input("ERROR_02:Invalid values in spreadsheet please contact: 612-703-9647")
                
            if(string1.find(item) == True):
                input("ERROR_02:Invalid values in spreadsheet please contact: 612-703-9647")
                
            if(string2.find(item) == True):
                input("ERROR_02:Invalid values in spreadsheet please contact: 612-703-9647")
                
            if(string3.find(item) == True):
                input("ERROR_02:Invalid values in spreadsheet please contact: 612-703-9647")
                
   
    for applicant in applicants:
        CreateApplicantPdf(str(applicant[0]),str(applicant[1]),str(applicant[2]),str(applicant[3]),TemplateFileName1,TemplateFileName2,path,date)


if __name__ == '__main__':
    try:
        print("If you are experiencing issues, please contact: 615-558-1000")
        print("Fadlan wac numbarkan: 615-558-1000, haddii aad ku aragto qalad. ")
        print("This application may take a few minutes, please wait.")
        Excel2Pdf("EmployeeList.xlsx","TimeSheetTemplate1.pdf","TimeSheetTemplate2.pdf" )
        input("Press enter to exit")
    except Exception as e:
        print("Error has occurred please contact:  615-558-1000 and tell them this error occurred: /n", e)
    
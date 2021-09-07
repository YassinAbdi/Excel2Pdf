# fill_by_overlay.py
import pdfrw
from reportlab.pdfgen import canvas
import os
from random import randint
import datetime
from PyPDF2 import PdfFileMerger
from datetime import date, timedelta


TrueStartingDate1 = datetime.datetime(2021, 4, 5)
TrueStartingDate2 = datetime.datetime(2021, 4, 12)

def random_with_N_digits(n):
    range_start = 10**(n-1)
    range_end = (10**n)-1
    return randint(range_start, range_end)

def write_PCA_OrgName(Canvas):
    hight1 = 680
    width1 = 40

    Canvas.drawString(width1, hight1, "Alpha Home Care Provider")

def write_PCA_OrgNumber(Canvas):
    hight1 = 680
    width1 = 490

    Canvas.drawString(width1, hight1, "612-315-4468")

def add_Time_Overlay1(canvas,startDate):
    hight1 = 645
    width1 = 140
    #write the Recipent information
    newDate = TrueStartingDate1
    for num in range(7):
        canvas.drawString(width1 + (num * 63) , hight1, newDate.strftime('%m/%d/%y'))
        newDate = newDate + timedelta(1)

def add_Time_Overlay2(canvas,startDate):
    hight1 = 645
    width1 = 140
    #write the Recipent information
    newDate = TrueStartingDate2
    for num in range(7):
        canvas.drawString(width1 + (num * 63) , hight1, newDate.strftime('%m/%d/%y'))
        newDate = newDate + timedelta(1)

def create_overlay(NameOfRecipent,IdOfMember,NameOfPCA,PcaNpi,FileName,startDate,secondWeek):
    """
    Create the data that will be overlayed on top
    of the form that we want to fill
    """
    #varibles that will be need for writing information
    hight1 = 125
    hight2 = 55

    width1 = 40
    width2 = 240

    c = canvas.Canvas(FileName)
    if(secondWeek == True):
        add_Time_Overlay1(c,startDate)
    else:
        add_Time_Overlay2(c,startDate)
    write_PCA_OrgName(c)
    write_PCA_OrgNumber(c)
    #write the Recipent information
    c.drawString(width1, hight1, NameOfRecipent)
    c.drawString(width2, hight1, IdOfMember)


    #write the PCA information
    c.drawString(width1, hight2, NameOfPCA)
    c.drawString(width2, hight2, PcaNpi)

    #save file
    c.save()

def merge_pdfs(form_pdf, overlay_pdf, output):
    """
    Merge the specified fillable form PDF with the 
    overlay PDF and save the output
    """
    form = pdfrw.PdfReader(form_pdf)
    olay = pdfrw.PdfReader(overlay_pdf)
    
    for form_page, overlay_page in zip(form.pages, olay.pages):
        merge_obj = pdfrw.PageMerge()
        overlay = merge_obj.add(overlay_page)[0]
        pdfrw.PageMerge(form_page).add(overlay).render()
        
    writer = pdfrw.PdfWriter()
    writer.write(output, form)

    #delete overlay pdf
    os.remove(overlay_pdf)

def concatenate_pdf(time_sheet1, time_sheet2, output):
    pdfs = [time_sheet1, time_sheet2]

    merger = PdfFileMerger()

    for pdf in pdfs:
        merger.append(pdf)

    merger.write(output)
    merger.close()
    #delete temp pdf
    os.remove(time_sheet1)
    os.remove(time_sheet2)

def CreateFolderForPdf(Date):
    #create a folder
    name = Date.strftime("%I_%M_%S_%p,%m_%d_%Y")
    if(os.path.isfile(name) == True):
        name = name + "("+str(random_with_N_digits(6))+")"
    try:
        os.mkdir(name)

    except:
        pass
    return name

def CreateApplicantPdf(NameOfRecipent,IdOfMember,NameOfPCA,PcaNpi,FormName, FormName2,Path,Date):
    filePrefix = str(Path) + "/"  
    FileNameDate = Date.strftime("%m_%d_%Y")
    fileExtention = ".pdf"
    pdfNumber = str(random_with_N_digits(6))
    TempPdfName1 = filePrefix+"Temp1_"+NameOfRecipent+"_"+pdfNumber+"_"+FileNameDate+fileExtention
    TempPdfName2 = filePrefix+"Temp2_"+NameOfRecipent+"_"+pdfNumber+"_"+FileNameDate+fileExtention
    TempFile1 = filePrefix+"Tempfile1.pdf"
    TempFile2 = filePrefix+"Tempfile2.pdf"
    PdfName = filePrefix+NameOfRecipent+"_"+pdfNumber+"_"+FileNameDate+fileExtention
    startDate = date.today()

    create_overlay(NameOfRecipent,IdOfMember,NameOfPCA,PcaNpi,TempPdfName1,startDate,True)
    merge_pdfs(FormName,TempPdfName1,TempFile1)

    create_overlay(NameOfRecipent,IdOfMember,NameOfPCA,PcaNpi,TempPdfName2,startDate,False)
    merge_pdfs(FormName2,TempPdfName2,TempFile2)

    concatenate_pdf(TempFile1,TempFile2,PdfName)

if __name__ == '__main__':
    TemplateName1 = 'TimeSheetTemplate1.pdf'
    TemplateName2 = 'TimeSheetTemplate2.pdf'
    NameOfRecipent = "Samakab ,Ahmed"
    IdOfMember = "05541636"
    NameOfPCA = "Ayan M Jama"
    PcaNpi = "A932638300"
    TempPdfName = 'simple_form_overlay.pdf'
    dateFormated = datetime.datetime.now().strftime("%m/%d/%Y")
    date = datetime.datetime.now()
    path = CreateFolderForPdf(date)
    CreateApplicantPdf(NameOfRecipent,IdOfMember,NameOfPCA,PcaNpi,TemplateName1,TemplateName2,path,date)

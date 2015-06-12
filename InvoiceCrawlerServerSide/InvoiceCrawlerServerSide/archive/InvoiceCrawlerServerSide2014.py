import clr
import time
import datetime
import os
clr.AddReference("System.Data")
from System.Data import *
from System.Data.SqlClient import *
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

def sqlInsert(sqlCommand):
    with SqlConnection(connectionString) as connection:
        command = SqlCommand(sqlCommand, connection)
        try:
            connection.Open()
            command.ExecuteNonQuery()
            return "success"
        except Exception, err:
            logWrite(logFilePath, str(err))
            return "error"
        finally:
            connection.Close()

def logWrite(fileLoc, textToWrite):
    with open(fileLoc, 'a') as file:
        currentDT = time.asctime(time.localtime(time.time()))
        file.write(str(currentDT) + " - " + textToWrite + "\n\n")

def isValidInvoice(path, fileName):
    isValid = True
    #Exclude file names with CR
    if fileName.lower().find("cr") != -1:
        isValid = False

    #Exclude file names with void
    if fileName.lower().find("void") != -1:
        isValid = False

    #Make sure this is an excel document
    if fileName.lower().find(".xlsx") == -1:
        isValid = False

    #Exclude Master files
    if fileName.lower().find("master") != -1:
        isValid = False
        
    #Exclude Temp files
    if fileName.lower().find("~$") != -1:
        isValid = False
        
    #Exclude Dave folder
    if path.lower().find("dave") != -1:
        isValid = False
        
    #Exclude PSI Upload folder
    if path.lower().find("uploaded") != -1:
        isValid = False

    return isValid

def validateExcelData(data):
    isValid = True
    if str(data['total']).lower() == "qoute":
        isValid = False

    if str(data['total']).lower() == "n/c":
        isValid = False

    if str(data['total']).lower() == "#n/a":
        isValid = False

    return isValid

def formatExcelData(data):
    if str(data['total']) == 'None':
        data['total'] = "0.00"

    if str(data['date']) == "None":
        data['date'] = "01/01/1901"

    if str(data['invoiceNumber']) == 'None':
        data['invoiceNumber'] = ""

    if str(data['pieces']) == 'None':
        data['pieces'] = 0
    else:
        data['pieces'] = 0

    if str(data['wieght']) == 'None':
        data['wieght'] = 0
    else:
        data['wieght'] = int(data['wieght'])

    data["filePath"] = data["filePath"].replace("'", "''")

    return data

def getFAC(path):
    if path.lower().find("atlanta") != -1:
        return "ATL"

    if path.lower().find("baltimore") != -1:
        return "BWI"

    if path.lower().find("hartford") != -1:
        return "BDL"

    if path.lower().find("angeles") != -1:
        return "LAX"

    if path.lower().find("miami") != -1:
        return "MIA"

    if path.lower().find("orlando") != -1:
        return "MCO"

    if path.lower().find("tampa") != -1:
        return "TPA"

    return ""

def getExcelData(pathToDoc, nplOrNpld):
    try:
        workbook = excel.Workbooks.Open(pathToDoc)
        worksheet = workbook.ActiveSheet
        data = {
            'total': worksheet.Cells(1,1).Value2,
            'date': convertExcelDT(worksheet.Cells(1,2).Value2),
            'invoiceNumber': worksheet.Cells(1,3).Value2,
            'pieces': worksheet.Cells(1,4).Value2,
            'wieght': worksheet.Cells(1,5).Value2
            }
        data['fac'] = getFAC(pathToDoc)
        data['weekNumber'] = getWeekNumber(data['invoiceNumber'])
        data['customerCode'] = data['invoiceNumber'][1:5]
        data['filePath'] = pathToDoc
        data['company'] = nplOrNpld
        data = formatExcelData(data)
            
        validData = validateExcelData(data)
        if validData == False:
            return "error"
        else:
            return data
    except Exception, err:
        print str(err)
        return "error"
    finally:
        workbook.Close(False)

def convertExcelDT(float):
    seconds = (float - 25569) * 86400.0
    converted = datetime.datetime.utcfromtimestamp(seconds)
    return converted

def getWeekNumber(invoiceNumber):
    invoiceNumber = invoiceNumber.replace(" ","")
    weekNumber = invoiceNumber[4:6]
    return weekNumber

def main(dirToSearch, company):
    for path, subdir, files in os.walk(dirToSearch):
        for fileName in files:
            invoicePath = path + "\\" + fileName
            valid = isValidInvoice(path, fileName)
            if valid == True:
                data = getExcelData(invoicePath, company)
                if data != "error":
                    sqlQuery = ("INSERT INTO Reporting.dbo.InvoiceCosts2014 VALUES("
                			    "'{0}',"
                			    "'{1}',"
                			    "'{2}',"
                			    "'{3}',"
                			    "'{4}',"
                			    "'{5}',"
                			    "'{6}',"
                			    "'{7}',"
                			    "'{8}',"
                			    "'{9}')").format(data['total'],
                			    data['date'],
                			    data['invoiceNumber'],
                			    data['pieces'],
                			    data['wieght'],
                			    data['customerCode'],
                			    data['fac'],
                			    data['filePath'],
                			    data['weekNumber'],
                			    data['company'])
                    sqlResaults = sqlInsert(sqlQuery)
                    print invoicePath
                    print ""
                    if sqlResaults == "error":
                        logWrite(logFilePath, sqlQuery + "\n***************\n\n")

#connectionString = "Data Source='NPLDM1'; Initial Catalog='Reporting'; User Id='invoiceCrawler'; Password='gmt-500RRC4'"
connectionString = "Server='192.168.0.201';User ID='invoiceCrawler';Password='gmt-500RRC4me';Database='Reporting'"
queryString = "SELECT * FROM Reporting.dbo.InvoiceCosts WHERE [Total] = '405.00'"
dirToSearch = "O:\\NPL Invoices\\2014 INVOICES"
dirToSearchNPLD = "O:\\NPL Dedicated Invoices"
logFilePath = "C:\\Users\\Shawn\\Desktop\\invoiceLog.txt"
sqlClearTable = "DELETE FROM Reporting.dbo.InvoiceCosts2014"
excel = Excel.ApplicationClass()
excel.Visible = False
excel.DisplayAlerts = False
        

sqlInsert(sqlClearTable)
main(dirToSearch, "NPL")
excel.Quit()
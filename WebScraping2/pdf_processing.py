import PyPDF2 as p2
import xlsxwriter

LwyEmial = []
LwyName = []
LwyStatus = []
LwyNumbers = []
LwyAddress = []

def getaddressline(document):
    filepath = str("/Users/michael4228/Desktop/PDFs/"+str(document)+".pdf")
    # print(filepath)
    PDFfile = open(filepath,'rb')
    pdfread = p2.PdfReader(PDFfile)
    x = pdfread.getPage(0)
    pagetext = str(x.extract_text())
    # print(pagetext)
    y = pagetext.splitlines()
    # print(y[4])
    return y[4]
def getalladdress():
    for i in range(1,1332):
        line = getaddressline(i)
        # print(line)
        LwyAddress.append(line)
getalladdress()
print(LwyAddress)


def getemailline(document):
    filepath = str("/Users/michael4228/Desktop/PDFs/"+str(document)+".pdf")
    # print(filepath)
    PDFfile = open(filepath,'rb')
    pdfread = p2.PdfReader(PDFfile)
    x = pdfread.getPage(0)
    pagetext = str(x.extract_text())
    # print(pagetext)
    y = pagetext.splitlines()
    # print(y[6])
    return y[6]
def getallemails():
    for i in range(1,1332):
        line = getemailline(i)
        # print(line)
        LwyEmial.append(line)
getallemails()
print(LwyEmial)

def getnmline(document):
    filepath = str("/Users/michael4228/Desktop/PDFs/"+str(document)+".pdf")
    # print(filepath)
    PDFfile = open(filepath,'rb')
    pdfread = p2.PdfReader(PDFfile)
    x = pdfread.getPage(0)
    pagetext = str(x.extract_text())
    # print(pagetext)
    y = pagetext.splitlines()
    # print(y[2])
    return y[2]
def getnms():
    for i in range(1,1332):
        line = getnmline(i)
        # print(line)
        LwyName.append(line)
getnms()
print(LwyName)

def getstatusline(document):
    filepath = str("/Users/michael4228/Desktop/PDFs/"+str(document)+".pdf")
    # print(filepath)
    PDFfile = open(filepath,'rb')
    pdfread = p2.PdfReader(PDFfile)
    x = pdfread.getPage(0)
    pagetext = str(x.extract_text())
    # print(pagetext)
    y = pagetext.splitlines()
    # print(y[3])
    return y[3]
def getsts():
    for i in range(1,1332):
        line = getstatusline(i)
        # print(line)
        LwyStatus.append(line)
getsts()
print(LwyStatus)

def getnumberline(document):
    filepath = str("/Users/michael4228/Desktop/PDFs/"+str(document)+".pdf")
    # print(filepath)
    PDFfile = open(filepath,'rb')
    pdfread = p2.PdfReader(PDFfile)
    x = pdfread.getPage(0)
    pagetext = str(x.extract_text())
    # print(pagetext)
    y = pagetext.splitlines()
    # print(y[5])
    return y[5]
def getnbs():
    for i in range(1,1332):
        line = getnumberline(i)
        # print(line)
        LwyNumbers.append(line)
getnbs()
print(LwyNumbers)



dcwb= xlsxwriter.Workbook('ws1.xlsx')
dcws = dcwb.add_worksheet('address')
col = 0
row = 0
for i in LwyAddress:
    dcws.write(row,0,i)
    row+=1
dcwb.close()

dcwb= xlsxwriter.Workbook('ws2.xlsx')
dcws = dcwb.add_worksheet('Name')
col = 0
row = 0
for i in LwyName:
    dcws.write(row,0,i)
    row+=1
dcwb.close()

dcwb= xlsxwriter.Workbook('ws4.xlsx')
dcws = dcwb.add_worksheet('Status')
col = 0
row = 0
for i in LwyStatus:
    dcws.write(row,0,i)
    row+=1
dcwb.close()

dcwb= xlsxwriter.Workbook('ws5.xlsx')
dcws = dcwb.add_worksheet('Phone')
col = 0
row = 0
for i in LwyNumbers:
    dcws.write(row,0,i)
    row+=1
dcwb.close()




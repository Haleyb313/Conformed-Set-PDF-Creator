import os
import pandas as pd
import openpyxl
from PyPDF2 import PdfFileMerger, PdfFileReader
import glob
import PyPDF4

# Goal:  Create one combined PDF out of all the individual PDFs located in respective folders
#       1) Open Excel file to find out 1) folder locations and 2) Combined PDF file names
#       2) Go into those folders, get all PDFS, save as one document using the designated name
#           a) include page labels
#           b) include page bookmarks
#       3) Do this for all 3 Volumes!

# Prior to running this code, several Macros within the Excel file are run.  These include:
#       1) Building a list of all Conformed Drawings (the most recent version of each drawing)
#       2) Getting a handle to each PDF, updating the name, and saving it within a designated folder

# To start, get a handle to each volume's Master PDF Excel list
#       While the file is renamed each week with a new date, we do know it is the only Excel file in the folder
#       So we can go in, get a list of all Excel files, then grab the 1st (and only) file

# Within each file, we created a table designating folder names and locations.
#       Make a dataframe of that information to reference later

# Volume 1 - Build A
pathVol1 = r"C:\Users\Documents\Conformed Comprehensive Drawings\Volume_1_BuildA"
extensionXLSM = 'xlsm'
os.chdir(pathVol1)
resultVol1 = glob.glob('*.{}'.format(extensionXLSM))

tableVol1 = pathVol1 + "\\" + resultVol1[0]

dfTableVol1 = pd.read_excel(tableVol1,
    sheet_name="ConformedDrawings",
    engine='openpyxl',
    index_col=False,
    names=["Folder", "Sub Folder", "Folder Path", "File Name"],
    usecols="F:H,K",
    header=8,
    nrows=19
)

# Volume 2 - Build B
pathVol2 = r"C:\Users\Documents\Conformed Comprehensive Drawings\Volume_2_BuildB"
extensionXLSM = 'xlsm'
os.chdir(pathVol2)
resultVol3 = glob.glob('*.{}'.format(extensionXLSM))

tableVol2 = pathVol2 + "\\" + resultVol3[0]

dfTableVol2 = pd.read_excel(tableVol2,
    sheet_name="ConformedDrawings",
    engine='openpyxl',
    index_col=False,
    names=["Folder", "Sub Folder", "Folder Path", "File Name"],
    usecols="F:H,K",
    header=8,
    nrows=6
)

# Volume 3 - Build C
pathVol3 = r"C:\Users\Documents\Conformed Comprehensive Drawings\Volume_3_BuildC"
extensionXLSM = 'xlsm'
os.chdir(pathVol3)
resultVol3 = glob.glob('*.{}'.format(extensionXLSM))

tableVol3 = pathVol3 + "\\" + resultVol3[0]

dfTableVol3 = pd.read_excel(tableVol3,
    sheet_name="ConformedDrawings",
    engine='openpyxl',
    index_col=False,
    names=["Folder", "Folder Path", "File Name"],
    usecols="G:H,J",
    header=8,
    nrows=1
)

# 2) Loop through all 3 of those dataframe tables:  For each row in Folder Path column...
#       a) Follow that Path into that folder
#       b) Get a list of all PDFs
#       c) Combine them
#       d) Save the single file using the name in the File Name column

from PyPDF2 import PdfFileWriter, PdfFileMerger, PdfFileReader
import PyPDF2.pdf as PDF

def appendPDFS(pathGrabPDFS, pdfsCount, finalPDFPath, finalPDFName):

        os.chdir(pathGrabPDFS)
        # pls holds all the pagelabels as we iterate through multiple pdfs
        pls = PDF.ArrayObject()
        # used to offset bookmarks
        pageCount = 0
        cpdf = PdfFileMerger()
        # pdffiles is a list of all files to be merged

        for i in range(len(pdfsCount)):
            pdfBookmark = str(pdfsCount[i]).replace('.pdf', '')
            output = PdfFileWriter()
            tmppdf = PdfFileReader(pdfsCount[i], 'rb')
            output.addBookmark(pdfBookmark, 0, parent=None)
            output.addPage(tmppdf.getPage(0))
            os.chdir(finalPDFPath)
            output.write(finalPDFName + ".pdf")
            os.chdir(pathGrabPDFS)

        for i in range(len(pdfsCount)):
            os.chdir(finalPDFPath)
            tmppdf = PdfFileReader(pdfsCount[i], 'rb')
            cpdf.append(tmppdf)
            # copy all the pagelabels which I assume is present in all files
            # you could use 'try' in case no pagelabels are present
            plstmp = tmppdf.trailer['/Root']['/PageLabels']['/Nums']
            # sometimes keys are indirect objects
            # so, iterate through each pagelabel and...
            for j in range(len(plstmp)):
                # ... get the actual values
                plstmp[j] = plstmp[j].getObject()
                # offset pagenumbers by current count of pages
                if isinstance(plstmp[j], int):
                    plstmp[j] = PDF.NumberObject(plstmp[j] + pageCount)
            # once all the pagelabels are processed I append to pls
            pls += plstmp
            #increment pageCount
            pageCount += tmppdf.getNumPages()

        pagenums = PDF.DictionaryObject()
        pagenums.update({PDF.NameObject('/Nums') : pls})
        pagelabels = PDF.DictionaryObject()
        pagelabels.update({PDF.NameObject('/PageLabels') : pagenums})
        cpdf.output._root_object.update(pagelabels)

        os.chdir(finalPDFPath)
        cpdf.write(finalPDFName+".pdf")


# Volume 1 - Build A
for i in dfTableVol1.itertuples():
    pathGrabPDFS = i[3] # Original file location, A1.105 - Schedules
    finalPDFName = i[4] # 07   A1.105 - Schedules
    finalPDFPath = pathVol1 # final file location

    os.chdir(pathGrabPDFS)
    extensionPDF = 'pdf'
    pdfsCount = glob.glob('*.{}'.format(extensionPDF))

    appendPDFS(pathGrabPDFS, pdfsCount, finalPDFPath, finalPDFName)

# Volume 2 - Build B
for i in dfTableVol2.itertuples():
    pathGrabPDFS = i[3] # Original file location, A1.105 - Schedules
    finalPDFName = i[4] # 07   A1.105 - Schedules
    finalPDFPath = pathVol2 # final file location

    os.chdir(pathGrabPDFS)
    extensionPDF = 'pdf'
    pdfsCount = glob.glob('*.{}'.format(extensionPDF))

    appendPDFS(pathGrabPDFS, pdfsCount, finalPDFPath, finalPDFName)

# Volume 3 - Build C
for i in dfTableVol3.itertuples():
    pathGrabPDFS = i[2] #Original file location, A1.105 - Schedules
    finalPDFName = i[3] #07   A1.105 - Schedules
    finalPDFPath = pathVol3 #final file location

    os.chdir(pathGrabPDFS)
    extensionPDF = 'pdf'
    pdfsCount = glob.glob('*.{}'.format(extensionPDF))

    appendPDFS(pathGrabPDFS, pdfsCount, finalPDFPath, finalPDFName)


def addBookmarks(pathGrabPDFS, pdfsCount, finalPDFPath, finalPDFName):
    os.chdir(pathGrabPDFS)
    # pls holds all the pagelabels as we iterate through multiple pdfs
    pls = PDF.ArrayObject()
    # used to offset bookmarks
    pageCount = 0

    for i in range(len(pdfsCount)):
        os.chdir(pathGrabPDFS)
        pdfBookmark = str(pdfsCount[i]).replace('.pdf', '')
        finalFile = finalPDFName + ".pdf"

        pdf_object = open(pdfsCount[i], "rb")
        output = PdfFileWriter()
        input = PdfFileReader(pdf_object)
        output.addPage(input.getPage(0))  # insert page
        output.addBookmark(pdfBookmark, 0, parent=None)  # add bookmark
        os.chdir(finalPDFPath)
        outputStream = open(finalFile, "wb")
        output.write(outputStream)
        outputStream.close()
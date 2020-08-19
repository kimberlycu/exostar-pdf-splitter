# Author: Kimberly Cu
# Date: 8-19-2020
# Project: Exostar PDF Splitter
# Description: Python script to split batch purchase order downloads from Exostar into separate PDF files
# with the order number and revision number as the file name.

from PyPDF2 import PdfFileWriter, PdfFileReader

# Converts a string of numbers and chars into a string with only numbers. Returns string.
def turnToNum(x):
  numeric_filter = filter(str.isdigit, x)
  numeric_string = "".join(numeric_filter)
  return numeric_string

# Extracts order and revision number from given page number. Returns order and revision number.
def orderRevNum(readerObj, currPageNum):
  pageObj = readerObj.getPage(currPageNum)
  text = pageObj.extractText().split(" ")
  orderNum = turnToNum(text[0])
  revNum = turnToNum(text[4])
  return orderNum, revNum

# Writes pages stored from PdfFileWriter into a PDF file with order and revision number as the file name.
# Closes the file stream.
def writeToFile(oNum, rNum, writerObj):
    output = oNum + " POC" + rNum + ".pdf"
    outputStream = open(output, "wb")
    writerObj.write(outputStream)
    outputStream.close()

# Initialize a reader and writer object.
inputStream = open("test3.pdf", "rb")
reader = PdfFileReader(inputStream)
writer = PdfFileWriter()

# Get order number and revision number from first page of PDF.
orderNum, revNum = orderRevNum(reader, 0)

# For each page in PDF, get order and revision number.
for pageNum in range(reader.getNumPages()):
  currOrderNum, currRevNum = orderRevNum(reader, pageNum)

# If order number of the last page does not match the current order number, write pages from writer object to
# a PDF with the order and revision number as the file name. Close output file stream. Reset writer object to get rid 
# of previously passed in pages. Set the current order and revision number as the older order and revision number
# for comparison.
  if orderNum != currOrderNum:
    writeToFile(orderNum, revNum, writer)
    writer = PdfFileWriter()
    orderNum = currOrderNum
    revNum = currRevNum

# Save current page to writer object.
  writer.addPage(reader.getPage(pageNum))

# If there are no more pages, write pages from writer object to new PDF with the order and revision number as the
# file name. Close output file stream.
  if pageNum == reader.getNumPages()-1:
    writeToFile(orderNum, revNum, writer)

# Close input file stream.
inputStream.close()

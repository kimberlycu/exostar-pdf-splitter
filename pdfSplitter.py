# Author: Kimberly Cu
# Date: 8-18-2020
# Project: Exostar PDF Splitter
# Description: Python script to split batch purchase order downloads from Exostar into separate PDF files
# with the order number and revision number as the file name.

from PyPDF2 import PdfFileWriter, PdfFileReader

# Converts a string of chars and numbers into a string with only numbers. Returns string.
def turnToNum(x):
  numeric_filter = filter(str.isdigit, x)
  numeric_string = "".join(numeric_filter)
  return numeric_string

# Initialize a reader and writer object.
inputStream = open("test2.pdf", "rb")
reader = PdfFileReader(inputStream)
writer = PdfFileWriter()

# Parse first page of the PDF with space delimiter. Extract order number and revision number.
pageObj = reader.getPage(0)
text = pageObj.extractText().split(" ")
orderNum = turnToNum(text[0])
revNum = turnToNum(text[4])

# For each page in PDF, extract order and revision number.
for pageNum in range(reader.getNumPages()):
  currPageObj = reader.getPage(pageNum)
  currText = currPageObj.extractText().split(" ")
  currOrderNum = turnToNum(currText[0])
  currRevNum = turnToNum(currText[4])

# If order number of the last page does not match the current order number, write pages from writer object to
# a PDF with the order and revision number as the file name. Close output file stream. Reset writer object to get rid 
# of previously passed in pages. Set the current order and revision number as the older order and revision number
# for comparison.
  if orderNum != currOrderNum:
    output = orderNum + " POC" + revNum + ".pdf"
    outputStream = open(output, "wb")
    writer.write(outputStream)
    outputStream.close()
    writer = PdfFileWriter()
    orderNum = currOrderNum
    revNum = currRevNum

# Save current page to writer object.
  writer.addPage(reader.getPage(pageNum))

# If there are no more pages, write pages from writer object to new PDF with the order and revision number as the
# file name. Close output file stream.
  if pageNum == reader.getNumPages()-1:
    output = orderNum + " POC" + revNum + ".pdf"
    outputStream = open(output, "wb")
    writer.write(outputStream)
    outputStream.close()

# Close input file stream.
inputStream.close()

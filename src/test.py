import os
import comtypes.client
import time


wdFormatPDF = 17
in_file=r'C:\Users\Pallavi A\projects\Olacabs1.doc'
out_file=r'C:\Users\Pallavi A\projects\Olacabs1.pdf'

# print out filenames
print(in_file)
print(out_file)


# create COM object
word = comtypes.client.CreateObject('Word.Application')
# key point 1: make word visible before open a new document
word.Visible = True
# key point 2: wait for the COM Server to prepare well.
time.sleep(3)

# convert docx file 1 to pdf file 1
doc=word.Documents.Open(in_file) # open docx file 1
doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
doc.Close() # close docx file 1
word.Visible = False
word.Quit() # close Word Application
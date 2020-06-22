import sys
import os
import comtypes.client
import time
import win32com.client


# This is when we're using this program in the cmd which isn't our case
# in_file = os.path.abspath(r"C:\Users\PC\PycharmProjects\file-converter\Canon In D by Pachelbel")
# out_file = os.path.abspath(r"C:\Users\PC\PycharmProjects\file-converter\Canon In D by Pachelbel")

wdFormatPDF = 17

in_file = "Canon In D by Pachelbel"
out_file = "Canon In D by Pachelbel"
# absolute path is needed
# be careful about the slash '\', use '\\' or '/' or raw string r"..."
in_file= r"C:\Users\PC\PycharmProjects\file-converter\Canon In D by Pachelbel.docx"
out_file= r"C:\Users\PC\PycharmProjects\file-converter\Canon In D by Pachelbel.pdf"



# create COM object
word = win32com.client.Dispatch('Word.Application')
# key point 1: make word visible before open a new document
word.Visible = True

# convert docx file 1 to pdf file 1
doc=word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Visible = False

word.Quit()
import sys
import os
import comtypes.client
import time
import win32com.client
import subprocess
from tkinter.filedialog import askopenfilename, Tk


def wordToPdf():
    Tk().withdraw() # Avoids showing tk window
    filename = askopenfilename()
    wdFormatPDF = 17
    in_file = filename
    out_file = filename[:-4]+".pdf"
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = True
    doc=word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Visible = False
    word.Quit()

def pdfToWord():
    Tk().withdraw() # Avoids showing tk window
    filename = askopenfilename()
    wdFormatPDF = 17
    in_file = filename
    out_file = filename[:-4]+".pdf"
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = True
    doc=word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Visible = False
    word.Quit()

def pdfCombiner():
    Tk().withdraw() # Avoids showing tk window
    filename = askopenfilename()
    wdFormatPDF = 17
    in_file = filename
    out_file = filename[:-4]+".pdf"
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = True
    doc=word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Visible = False
    word.Quit()


while True:
    print("""
        1. Word to pdf
        2. Pdf to word
        3. Combine pdfs
        4. Exit\n
    """)
    inp = str(input("What do you want to do?\n"))
    wordToPdf()
    print("word?")
    # if inp == "4":
    #     sys.exit(0)
    # elif inp == "1":
    #     wordToPdf()
    # elif inp == "2":
    #     pdfToWord()
    # elif inp == "3":
    #     pdfCombiner()
    # else:
    #     continue

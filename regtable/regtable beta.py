from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import os, sys, re, csv
from pathlib import Path

import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

from pikepdf import Pdf as pike

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator
from pdfminer.high_level import extract_text
import fnmatch

from difflib import SequenceMatcher


global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('240x130+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global SLbl
        Lbl = Label(text="Директория с папками:", background="white")
        Lbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=20)
        InputDirEntry.place(x=20, y=32)
        
        global InputDirBtn
        InputDirBtn = Button(text='Выбор', command=SelectDir)
        InputDirBtn.place(x=150, y=31, height=20)
        
        global StartBtn
        StartBtn = Button(text='Run', command=Run)
        StartBtn.place(x=20, y=55, height=20, width=45)



def SelectDir():
    global InputDir
    global InputFilesArray
    InputFilesArray = []

    InputDir = ""
    InputDir = filedialog.askdirectory(title='Выберите папку на обработку')
    if InputDir:
        InputDirEntry.configure(state = NORMAL)
        InputDirEntry.delete(0,END)
        InputDirEntry.insert(0,str(InputDir))
        InputDirEntry.configure(state = DISABLED)
        
        print('IDC: InputDir : {0}'.format(InputDir))
        print('*********************************')
        
        InputFilesArray.clear()
        for file in os.listdir(InputDir):
            if file.endswith(".pdf"):
                InputFilesArray.append(os.path.join(InputDir, file))
                print('**** File : {0}'.format(file))
                
    else:
        print('IDC: InputFile not selected')
        

def Run():
    for f in range(len(InputFilesArray)):
        print("")
        print("")
        print("**********************************************************")
        print('Документ №{0}: {1}'.format(f+1, Path(InputFilesArray[f]).name))
        
        #invoicedata = PlaceFileTextSearch(InputFilesArray[f])
        
        #PdfTextExtractor(InputFilesArray[f])
        DocNumberTextSearch(InputFilesArray[f])
        DocAddressTextSearch(InputFilesArray[f])
        #DocInnKppTextSearch(InputFilesArray[f])
        
        #if isinstance(invoicedata, list):
        #    print('Обработка {0} из {1}'.format(f, len(InputFilesArray)))
        #    print('ИНН, КПП документа: {0}'.format(invoicedata))


def PdfTextExtractor(file):

    with open(file, 'rb') as pdftomine:
        manager = PDFResourceManager()
        laparams = LAParams()
        dev = PDFPageAggregator(manager, laparams=laparams)
        interpreter = PDFPageInterpreter(manager, dev)
        pages = PDFPage.get_pages(pdftomine)

        for pagenumber, page in enumerate(pages):
            if pagenumber == 0:
                interpreter.process_page(page)
                layout = dev.get_result()
                
                for textbox in layout:
                    if isinstance(textbox, LTText):
                        for line in textbox:
                            text = line.get_text().replace('\n', '')
                            print('Text: {0}'.format(text))


def DocNumberTextSearch(file):

    with open(file, 'rb') as pdftomine:
        manager = PDFResourceManager()
        laparams = LAParams()
        dev = PDFPageAggregator(manager, laparams=laparams)
        interpreter = PDFPageInterpreter(manager, dev)
        pages = PDFPage.get_pages(pdftomine)

        for pagenumber, page in enumerate(pages):
            if pagenumber == 0:
                interpreter.process_page(page)
                layout = dev.get_result()
                
                for textbox in layout:
                    if isinstance(textbox, LTText):
                        for line in textbox:
                            text = line.get_text().replace('\n', '')
                            firstletter = text[0]
                            if firstletter == 'R':
                                if len(text) >= 20:
                                    print('============================================')
                                    print('Number: {0}'.format(text))
                                    return text
                                    break
                            

def DocAddressTextSearch(file):

    addressclfound = False
    addressclcounter = 0

    with open(file, 'rb') as pdftomine:
        manager = PDFResourceManager()
        laparams = LAParams()
        dev = PDFPageAggregator(manager, laparams=laparams)
        interpreter = PDFPageInterpreter(manager, dev)
        pages = PDFPage.get_pages(pdftomine)

        for pagenumber, page in enumerate(pages):
            if pagenumber == 0:
                interpreter.process_page(page)
                layout = dev.get_result()
                
                for textbox in layout:
                    if isinstance(textbox, LTText):
                        for line in textbox:
                            text = line.get_text().replace('\n', '')
                            
                            if addressclfound == True:
                                addressclcounter = addressclcounter + 1
                                if addressclcounter == 2:
                                    print('Address: {0}'.format(text))
                                    print('============================================')
                                    return text
                                    break
                            
                            if text[0] == '№' and text[2] == 'п' and text[3] == '/' and text[4] == 'п':
                                addressclfound = True
                                                        

def DocInnKppTextSearch(file):
    
    vendordatafound = False
    customerdatafound = False
    
    vendorinn = 'нет данных'
    vendorkpp = 'нет данных'
    customerinn = 'нет данных'
    customerkpp = 'нет данных'

    with open(file, 'rb') as pdftomine:
        manager = PDFResourceManager()
        laparams = LAParams()
        dev = PDFPageAggregator(manager, laparams=laparams)
        interpreter = PDFPageInterpreter(manager, dev)
        pages = PDFPage.get_pages(pdftomine)

        for pagenumber, page in enumerate(pages):
            if pagenumber == 0:
                interpreter.process_page(page)
                layout = dev.get_result()
                
                for textbox in layout:
                    if isinstance(textbox, LTText):
                        for line in textbox:
                            text = line.get_text().replace('\n', '')
                            
                            if vendordatafound == False:
                                if len(text) == 22 or len(text) == 20:
                                    invoiceinn = re.sub("[^0-9]", "", (text.partition("/")[0]))
                                    invoicekpp = re.sub("[^0-9]", "", (text.partition("/")[2]))
                                    if invoiceinn.isnumeric() and invoicekpp.isnumeric():
                                        if len(invoiceinn)==10 and len(invoicekpp)==9:
                                            vendorinn = invoiceinn
                                            vendorkpp = invoicekpp
                                            print('============================================')
                                            print('VENDOR DATA - ИНН: {0}, КПП: {1}'.format(vendorinn, vendorkpp))
                                            vendordatafound = True
                            else:
                                if customerdatafound == False:
                                    if len(text) == 24 or len(text) == 22 or len(text) == 20 or len(text) == 14:
                                        invoiceinn = re.sub("[^0-9]", "", (text.partition("/")[0]))
                                        invoicekpp = re.sub("[^0-9]", "", (text.partition("/")[2]))
                                        if invoiceinn.isnumeric():
                                            if len(invoiceinn) == 10 or len(invoiceinn) == 12:
                                                customerinn = invoiceinn
                                                if invoicekpp.isnumeric():
                                                    customerkpp = invoicekpp
                                                else:
                                                    customerkpp = 'null'
                                            customerdatafound = True
                                            print('CUSTOMER DATA - ИНН: {0}, КПП: {1}'.format(customerinn, customerkpp))
                                            print('============================================')
                                else:
                                    invoicedata = [customerinn, customerkpp]
                                    return invoicedata
                                    break
        return "NONE"


if __name__ == '__main__':
    main()

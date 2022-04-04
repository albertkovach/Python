from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import win32api
import win32print
import PyPDF2
global root
import os
import re


from io import StringIO

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser



def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = root.winfo_screenwidth()
    scrnh = root.winfo_screenheight()
    scrnw = scrnw//2
    scrnh = scrnh//2
    scrnw = scrnw - 185
    scrnh = scrnh - 150
    root.geometry('300x90+{}+{}'.format(scrnw, scrnh))
        
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
        global TextLbl
        TextLbl = Label(text="", background="white")
        TextLbl.place(x=20, y=45)
        
        global ChooseFileBtn
        ChooseFileBtn = Button(text='Выбрать', command=BtnCmd1)
        ChooseFileBtn.place(x=20, y=20, height=20)

        global RunBtn
        RunBtn = Button(text='Запуск', command=BtnCmd2)
        RunBtn.place(x=90, y=20, height=20)
        
        global SearchTextEntry
        SearchTextEntry = Entry(fg="black", bg="white", width=20)
        SearchTextEntry.place(x=150, y=20)
        
        global ShowAllBtn
        ShowAllBtn = Button(text='\/', command=Miner)
        ShowAllBtn.place(x=253, y=45, height=20)



def BtnCmd1():
    global filename
    filename = filedialog.askopenfilename(filetypes=(('', 'pdf'),))
    TextLbl.config(text = str(filename))
    print('Selected:', filename)


def BtnCmd2():
    global filename

    pdf = PyPDF2.PdfFileReader(filename)
    NumPages = pdf.getNumPages()
    SearchText = SearchTextEntry.get()
    #SearchText = "1"

    if SearchText:
        for i in range(0, NumPages):
            PageObj = pdf.getPage(i)
            print("this is page " + str(i)) 
            Text = PageObj.extractText() 
            ResSearch = re.search(SearchText, Text)
            print(ResSearch)
    
    print('End of search')

    
def BtnCmd3():
    global filename

    pdf = PyPDF2.PdfFileReader(filename)
    NumPages = pdf.getNumPages()

    for i in range(0, NumPages):
        PageObj = pdf.getPage(i)
        print("=========="+"this is page "+str(i)+"==========") 
        Text = PageObj.extractText() 
        print(Text)
    print('End of show')
    
    
def Miner():
    global filename
    rawtext = convert_pdf_to_string(filename)
    print(rawtext)
    print('End Miner')




def convert_pdf_to_string(file_path):
	output_string = StringIO()
	with open(file_path, 'rb') as in_file:
	    parser = PDFParser(in_file)
	    doc = PDFDocument(parser)
	    rsrcmgr = PDFResourceManager()
	    device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
	    interpreter = PDFPageInterpreter(rsrcmgr, device)
	    for page in PDFPage.create_pages(doc):
	        interpreter.process_page(page)
    #rawtext = output_string.getvalue()
	return(output_string.getvalue())


if __name__ == '__main__':
    main()

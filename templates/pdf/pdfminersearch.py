from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import win32api
import win32print
import PyPDF2
global root
import os

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator

from difflib import SequenceMatcher

global FirstPagesArray




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
        ShowAllBtn = Button(text='\/', command=BtnCmd3)
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
    Miner()
    ArrayInterpreter()
 
    
    
    
    
    
def Miner():
    global filename
    global FirstPagesArray
    print('============ Miner started ===========')
    
    word = 'Счет-фактура №'
    pagecounter = 1
    finded = False
    FirstPagesArray = []
    FirstPagesArray.clear

    pdftomine = open(filename, 'rb')
    manager = PDFResourceManager()
    laparams = LAParams()
    dev = PDFPageAggregator(manager, laparams=laparams)
    interpreter = PDFPageInterpreter(manager, dev)
    pages = PDFPage.get_pages(pdftomine)  

    for page in pages:
        interpreter.process_page(page)
        layout = dev.get_result()
        for textbox in layout:
            if isinstance(textbox, LTText):
                for line in textbox:
                    text = line.get_text()
                    similarity = similar(text, word)
                    if similarity > 0.9:
                        FirstPagesArray.append(pagecounter)
                        print('finded! page ' + str(pagecounter))
                        finded = True
        if finded:
            finded = False
        else:
            print('page ' + str(pagecounter))
        pagecounter = pagecounter + 1
        
    print('============ Miner ended =============')
    print('')




def ArrayInterpreter():
    print('====== ArrayInterpreter started ======')
    
    origpdf = PyPDF2.PdfFileReader(filename)
    NumPages = origpdf.getNumPages()
    print('NumPages: ' + str(NumPages))
    
    pdf_writer1 = PdfFileWriter()
    
    for k in range(len(FirstPagesArray)):
        if k+1 < len(FirstPagesArray):
            print('==========')
            print(str(FirstPagesArray[k]))
            print(str(FirstPagesArray[k+1]))
            
            print('----')
            for x in range(FirstPagesArray[k], FirstPagesArray[k+1]):
                print(x)
            print('----')
            print('')
            
    print('== Last ==')
    for x in range(FirstPagesArray[len(FirstPagesArray)-1], NumPages+1):
        print(x)

    print('====== ArrayInterpreter ended ======')
    print('')








def SavenameGenerate(Nameless, AddDate, Name, Part):
    global CreateDocTime
    
    if Nameless:
        if AddDate:
            if Part < 1:
                savename = ("Merged "+ CreateDocTime)
            else:
                savename = ("Merged, part " + str(Part) + " - " + CreateDocTime)
        else:
            if Part < 1:
                savename = ("Merged")
            else:
                savename = ("Merged, part " + str(Part))
    else:
        if AddDate:
            if Part < 1:
                savename = (Name + " " + CreateDocTime)
            else:
                savename = (Name + " " + str(Part) + " - " + CreateDocTime)
        else:
            if Part < 1:
                savename = (Name)
            else:
                savename = (Name + ", part " + str(Part))

    return savename




def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

if __name__ == '__main__':
    main()

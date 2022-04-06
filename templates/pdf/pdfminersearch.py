from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import win32api
import win32print
import PyPDF2
from pathlib import Path
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
    print('BtnCmd2:', filename)
    originalpdf = PyPDF2.PdfFileReader(filename)
    
    outputfile = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Divider', 'aaa.pdf')
    outputfilepath = Path(outputfile)
    
    pdf_writer = PyPDF2.PdfFileWriter()
    
    pdf_writer.addPage(originalpdf.getPage(1))
    pdf_writer.addPage(originalpdf.getPage(2))
    pdf_writer.addPage(originalpdf.getPage(3))
    pdf_writer.addPage(originalpdf.getPage(4))
    pdf_writer.addPage(originalpdf.getPage(5))
    
    pdf_writer.write(open(outputfile, 'wb'))
    pdf_writer = ''
    
    
    outputfile = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Divider', 'bbb.pdf')
    outputfilepath = Path(outputfile)
    
    pdf_writer = PyPDF2.PdfFileWriter()
    
    pdf_writer.addPage(originalpdf.getPage(6))
    pdf_writer.addPage(originalpdf.getPage(7))
    pdf_writer.addPage(originalpdf.getPage(8))
    pdf_writer.addPage(originalpdf.getPage(9))
    #pdf_writer.addPage(originalpdf.getPage(10))
    
    pdf_writer.write(open(outputfile, 'wb'))
    pdf_writer = ''

    
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
    
    pdftomine = ''
    manager = ''
    laparams = ''
    dev = ''
    interpreter = ''
    pages = ''




def ArrayInterpreter():
    print('====== ArrayInterpreter started ======')
    
    originalpdf = PyPDF2.PdfFileReader(filename)
    NumPages = originalpdf.getNumPages()

    
    print('NumPages: ' + str(NumPages))
    print('')
    
    savedir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Merger') 
    
    
    for k in range(len(FirstPagesArray)):
        if k+1 < len(FirstPagesArray):
        
            print('**** Документ №: ', str(k+1))
            outputfileop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Divider', 'onepage', (str(FirstPagesArray[k])+'.pdf'))
            outputfilemp = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Divider', 'multipage', (str(FirstPagesArray[k])+'.pdf'))
            print("Номер первой стр: {0}, номер сл.первой {1}".format(FirstPagesArray[k],FirstPagesArray[k+1]))
            #print("Итоговый файл: {0}".format(outputfile))
            print('Список страниц документа:')
            temparray = []
            temparray.clear()
            for x in range(FirstPagesArray[k], FirstPagesArray[k+1]):
                print(x)
                temparray.append(x)
                
            
            pdf_writer = PyPDF2.PdfFileWriter()
            
            for x in range(len(temparray)):
                pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))
                
                
            if len(temparray) == 1:
                pdf_writer.write(open(outputfileop, 'wb'))            
            elif len(temparray) % 2 == 1:
                _, _, w, h = originalpdf.getPage(0)['/MediaBox']
                pdf_writer.addBlankPage(w, h)
                pdf_writer.write(open(outputfilemp, 'wb'))
            else:
                pdf_writer.write(open(outputfilemp, 'wb'))
                
            pdf_writer = ''
            print('Обработка завершена')
            print('*******************')
            print('')


    print('**** Последний документ: ', str(len(FirstPagesArray)))
    outputfileop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Divider', 'onepage', (str(FirstPagesArray[len(FirstPagesArray)-1])+'.pdf'))
    outputfilemp = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'Divider', 'multipage', (str(FirstPagesArray[len(FirstPagesArray)-1])+'.pdf'))

    #print("Итоговый файл: {0}".format(outputfile))
    print('Список страниц документа:')
    
    temparray.clear()
    for x in range(FirstPagesArray[len(FirstPagesArray)-1], NumPages+1):
        print(x)
        temparray.append(x)
        
    pdf_writer = PyPDF2.PdfFileWriter()
    for x in range(len(temparray)):
        pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))


    if len(temparray) == 1:
        pdf_writer.write(open(outputfileop, 'wb'))            
    elif len(temparray) % 2 == 1:
        _, _, w, h = originalpdf.getPage(0)['/MediaBox']
        pdf_writer.addBlankPage(w, h)
        pdf_writer.write(open(outputfilemp, 'wb'))
    else:
        pdf_writer.write(open(outputfilemp, 'wb'))


    pdf_writer = ''
    print('Обработка завершена')
    print('****************************')


    print('')
    print(' Documents on file: ' + str(len(FirstPagesArray)))
    print('')
    print('====== ArrayInterpreter ended ======')
    print('')


#outputfileop = Path(DividerOutputDir, 'onepage', (str(FirstPagesArray[k])+'.pdf'))
#outputfilemp = Path(DividerOutputDir, 'onepage', (str(FirstPagesArray[k])+'.pdf'))

def make_even_page(in_fpath, out_fpath):
    reader = PyPDF2.PdfFileReader(in_fpath)
    writer = PyPDF2.PdfFileWriter()
    for i in range(reader.getNumPages()):
        writer.addPage(reader.getPage(i))
    if reader.getNumPages() % 2 == 1:
        _, _, w, h = reader.getPage(0)['/MediaBox']
        writer.addBlankPage(w, h)
    with open(out_fpath, 'wb') as fd:
        writer.write(fd)




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

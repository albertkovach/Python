import sys
import os
import ctypes



from pathlib import Path
import shutil

import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator

from difflib import SequenceMatcher



    




def main():

    # Disable cmd quick edit mode
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), 128)
    
    # Setup cmd window
    #os.system('color 3f')
    os.system('mode con: cols=73 lines=10')
    

    print("Обработка файла 1 из 1: 383.pdf")
    DivideMiner()



def DivideMiner():
    global DivideInputFile
    global DivideAllPagesCount
    global DivideFirstPagesArray
    
    DivideInputFile = Path("C:\\Users\\ORB User\\Desktop\\61.pdf").as_posix()
    DivideAllPagesCount = PDFCountPages(DivideInputFile)
    
    
    word = 'Счет-фактура №'
    pagecounter = 1
    finded = False
    DivideFirstPagesArray = []
    DivideFirstPagesArray.clear

    pdftomine = open(DivideInputFile, 'rb')
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
                    similarity = Similar(text, word)
                    if similarity > 0.9:
                        DivideFirstPagesArray.append(pagecounter)
                        finded = True
        if finded:
            finded = False
            
        sys.stdout.write(' \r\033[K')
        print("Поиск первых страниц... Чтение {0} из {1}, найдено: {2}".format(pagecounter, DivideAllPagesCount, len(DivideFirstPagesArray)), end='\r')
        pagecounter = pagecounter + 1
    
    pdftomine = ''
    manager = ''
    laparams = ''
    dev = ''
    interpreter = ''
    pages = ''
    
    if len(DivideFirstPagesArray) > 0:
        sys.stdout.write(' \r\033[K')
        print("Поиск завершен ! Найдено счет-фактур: {0}".format(len(DivideFirstPagesArray)), end='\r')
        #DivideFileMaker()
    else:
        sys.stdout.write(' \r\033[K')
        print("Поиск завершен, счет-фактур не найдено !", end='\r')
    close








def DivideFileMaker():
    global DivideInputFile
    global DivideAllPagesCount
    global DivideFirstPagesArray
    
    global DivideOutputDir
    
    global DivideIsRunning
    
    
    print('')
    print('Divide: FM: Started !')
    originalpdf = PyPDF2.PdfFileReader(DivideInputFile)
   
    for k in range(len(DivideFirstPagesArray)):
        if k+1 < len(DivideFirstPagesArray):
        
            print('**** Документ №: ', str(k+1))
            outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(DivideFirstPagesArray[k])+'.pdf'))
            print("Номер первой стр: {0}, номер сл.первой {1}".format(DivideFirstPagesArray[k],DivideFirstPagesArray[k+1]))
            print("Итоговый файл: {0}".format(outputfile))
            print('Список страниц документа:')
            
            temparray = []
            temparray.clear()
            for x in range(DivideFirstPagesArray[k], DivideFirstPagesArray[k+1]):
                print(x)
                temparray.append(x)
            
            pdf_writer = PyPDF2.PdfFileWriter()
            for x in range(len(temparray)):
                pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

            progresslbltxt = "Создание документов... Обработка {0} из {1}".format(k, len(DivideFirstPagesArray))
            DivideStatusLbl.config(text = progresslbltxt)

            pdf_writer.write(open(outputfile, 'wb'))
            pdf_writer = ''

            print('Обработка завершена')
            print('*******************')
            print('')


    print('**** Последний документ: ', str(len(DivideFirstPagesArray)))
    outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(DivideFirstPagesArray[k])+'.pdf'))
    print("Итоговый файл: {0}".format(outputfile))
    print('Список страниц документа:')
    
    temparray.clear()
    for x in range(DivideFirstPagesArray[len(DivideFirstPagesArray)-1], DivideAllPagesCount+1):
        print(x)
        temparray.append(x)
        
    pdf_writer = PyPDF2.PdfFileWriter()
    for x in range(len(temparray)):
        pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

    progresslbltxt = "Создание документов... Обработка {0} из {1}".format(len(DivideFirstPagesArray), len(DivideFirstPagesArray))
    DivideStatusLbl.config(text = progresslbltxt)

    pdf_writer.write(open(outputfile, 'wb'))
    pdf_writer = ''
    
    progresslbltxt = "Обработка файла завершена! Извлечено документов: {0}".format(len(DivideFirstPagesArray))
    DivideStatusLbl.config(text = progresslbltxt)

    print('Обработка завершена')
    print('****************************')

    DivideIsRunning = False
    DivideBlockGUI(False)
    
    print('Documents in file: ' + str(len(DivideFirstPagesArray)))
    print('Divide: FM: Ended !')
    
    msgbxlbl = 'Обработка файла завершена!'
    messagebox.showinfo("", msgbxlbl)
    
    
def Similar(a, b):
    return SequenceMatcher(None, a, b).ratio()



def PDFCountPages(inputfile):
    pdf = PdfFileReader(inputfile)
    pagecount = pdf.getNumPages()
    pdf = ""
    return pagecount



if __name__ == "__main__":
    main()

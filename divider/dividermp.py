from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import ctypes as ct

from threading import Thread
from multiprocessing import Process, Queue, current_process
import psutil

import time
import datetime as datetime2
from datetime import datetime

from pathlib import Path
import shutil
import os
import sys

import PyPDF2
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter
from PyPDF2 import PdfFileMerger

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator

from difflib import SequenceMatcher




# >>>>> GUI variables <<<<<
# -------------------------
global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

global CurrentUIpage
# -------------------------



# >>> Divide variables <<<
# -------------------------
global DivideInputFilesArray
global DivideInputFile
global DivideIsInputSel
global DivideReadyToGo

global DivideOutputDir
global DivideIsOutputSel

global DivideIsRunning
global DivideStartedTime
global DivideGUIisResized
global MaxProcessCount
# -------------------------



def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    datafile = "icon.ico"
    if not hasattr(sys, "frozen"):
        datafile = os.path.join(os.path.dirname(__file__), datafile)
    else:
        datafile = os.path.join(sys.prefix, datafile)
    root.iconbitmap(default=ResourcePath(datafile))
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()



class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("Divider")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        # >> Divide var init <<
        # ----------------------
        global DivideInputFilesArray
        global DivideInputFile
        global DivideOutputDir
        DivideInputFilesArray = []
        DivideInputFile = ""  
        DivideOutputDir = ""  
                
        global DivideIsInputSel
        global DivideIsOutputSel
        global DivideIsRunning
        global DivideReadyToGo
        global DivideGUIisResized
        DivideIsInputSel = False
        DivideIsOutputSel = False
        DivideIsRunning = False
        DivideReadyToGo = False
        DivideGUIisResized = False
        
        global DivideFirstPagesArray
        DivideFirstPagesArray = []
        # ----------------------
    
        canvas = Canvas(self, bg='white')
        canvas.create_line(15, 225, 360, 225)
        canvas.pack(fill=BOTH, expand=1)
    
        # >>>>>>>>>>>>>>> Mode header <<<<<<<<<<<<<<<
        # ===========================================
        global MainModeLbl
        MainModeLbl = Label(text="Разделение по коробам", background="white", font=("Arial", 12))
        MainModeLbl.place(x=16, y=6)







                ######## Divide widgets #######
                ################################
                
        # >>>>> Divide input file widgets block <<<<
        # ===========================================
        
        global DivideInputDirLbl
        DivideInputDirLbl = Label(text="Выберите файлы для разделения:", background="white", font=("Arial", 10))
        
        global DivideInputDirEntry
        DivideInputDirEntry = Entry(fg="black", bg="white", width=46)
        DivideInputDirEntry.configure(state = DISABLED)
        
        global DivideInputDirChooseBtn
        DivideInputDirChooseBtn = Button(text='Выбор', command=DivideInputDirChoose)
        
        global DivideInputDirPCountLbl
        DivideInputDirPCountLbl = Label(text="", background="white")
        
        
        # >>>>>>> Divide output widgets block <<<<<<
        # ===========================================
        
        global DivideOutputDirBtnLbl
        DivideOutputDirBtnLbl = Label(text="Папка для разделенных счет-фактур:", background="white", font=("Arial", 10))
        
        global DivideOutputDirBtn
        DivideOutputDirBtn = Button(text="Выбор", command=DivideOutputDirChoose)

        global DivideOutputDirEntry
        DivideOutputDirEntry = Entry(fg="black", bg="white", width=46)
        DivideOutputDirEntry.configure(state = DISABLED)
        
        
        global DivideStartDivisionBtn
        DivideStartDivisionBtn = Button(text='Выполнить разделение', command=DivideStartDivision)
        DivideStartDivisionBtn.configure(state = DISABLED)
        

        global DivideStatusLbl
        DivideStatusLbl = Label(text="", background="white")
        
        global DivideTimeLbl
        DivideTimeLbl = Label(text="", background="white")
        
        
        global DivideResizeGUIBtn
        DivideResizeGUIBtn = Button(text="➘", command=DivideResizeGUI)
        
        
        global DivideProcess1StatusLbl
        DivideProcess1StatusLbl = Label(text="Поток 1...", background="white", justify=LEFT)
        
        global DivideProcess2StatusLbl
        DivideProcess2StatusLbl = Label(text="Поток 2...", background="white", justify=LEFT)
        
        global DivideProcess3StatusLbl
        DivideProcess3StatusLbl = Label(text="Поток 3...", background="white", justify=LEFT)
        
        global DivideProcess4StatusLbl
        DivideProcess4StatusLbl = Label(text="Поток 4...", background="white", justify=LEFT)
        
        global DivideProcess5StatusLbl
        DivideProcess5StatusLbl = Label(text="Поток 5...", background="white", justify=LEFT)
        
        global DivideProcess6StatusLbl
        DivideProcess6StatusLbl = Label(text="Поток 6...", background="white", justify=LEFT)
        
        
        
        MainModeLbl.config(text = 'Разделение по счет-фактурам:')
    
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x225+{}+{}'.format(scrnw, scrnh))
        
        
        DivideInputDirLbl.place        (x=16, y=7+40)
        DivideInputDirEntry.place      (x=20, y=30+40)
        DivideInputDirChooseBtn.place  (x=305, y=29+40, height=20)
        DivideInputDirPCountLbl.place  (x=16, y=49+40)
        DivideOutputDirBtnLbl.place    (x=17, y=87+35)
        DivideOutputDirBtn.place       (x=20, y=110+35, height=20)
        DivideOutputDirEntry.place     (x=75, y=111+35)
        DivideStartDivisionBtn.place   (x=20, y=145+37, width=150, height=25)
        
        DivideStatusLbl.place          (x=175, y=184)
        DivideTimeLbl.place            (x=285, y=200)
        DivideResizeGUIBtn.place       (x=340, y=202, height=16)
        
        DivideProcess1StatusLbl.place     (x=20, y=230)
        DivideProcess2StatusLbl.place     (x=20, y=270)
        DivideProcess3StatusLbl.place     (x=20, y=310)
        DivideProcess4StatusLbl.place     (x=20, y=350)
        DivideProcess5StatusLbl.place     (x=20, y=390)
        DivideProcess6StatusLbl.place     (x=20, y=430)
        
        
        


####### Divide functions
#########################

def DivideInputFileChoose():
    global DivideInputFile
    global DivideIsInputSel
    global DivideAllPagesCount

    DivideInputFile = filedialog.askopenfilename(title='Выберите папку на обработку', filetypes=(('PDF document', 'pdf'),))
    if DivideInputFile:
        DivideInputDirEntry.configure(state = NORMAL)
        DivideInputDirEntry.delete(0,END)
        DivideInputDirEntry.insert(0,str(DivideInputFile))
        DivideInputDirEntry.configure(state = DISABLED)
        print('Divide: IFC: DivideInputFile :', DivideInputFile)
        DivideAllPagesCount = PDFCountPages(DivideInputFile)
        DivideInputDirPCountLbl.config(text = 'Страниц в файле: ' + str(DivideAllPagesCount))
        
        DivideIsInputSel = True
        DivideOutputDirCheck()
        DivideCheckIfReady()
    else:
        print('Divide: IFC: DivideInputFile not selected')


def DivideInputDirChoose():
    global DivideInputFilesArray
    global DivideIsInputSel
    global DivideAllPagesCount

    divideinputdir = filedialog.askdirectory(title='Выберите папку с файлами на обработку')
    if divideinputdir:
        DivideInputFilesArray = []
        ScanFolder(divideinputdir, ".pdf", DivideInputFilesArray)
        
        if len(DivideInputFilesArray) > 0:
            DivideInputDirEntry.configure(state = NORMAL)
            DivideInputDirEntry.delete(0,END)
            DivideInputDirEntry.insert(0,str(Path(divideinputdir).name))
            DivideInputDirEntry.configure(state = DISABLED)
            
            print('Divide: IFC: divideinputdir :', divideinputdir)
            
            counterthread = Thread(target=DividerPageCountThread)
            counterthread.start()
            
            DivideIsInputSel = True
            DivideOutputDirCheck()
            DivideCheckIfReady()
        else:
            DivideInputDirPCountLbl.config(text = "В папке нет файлов PDF !")
            DivideIsInputSel = False
            DivideOutputDirCheck()
            DivideCheckIfReady()
    else:
        print('Divide: IFC: DivideInputFile not selected')


def DividerPageCountThread():
    global DivideInputFilesArray
    #MainModeBackBtn.configure(state = DISABLED)
    
    allpagecount = 0
    for i in range (len(DivideInputFilesArray)):
        allpagecount = allpagecount + PDFCountPages(DivideInputFilesArray[i])
    
        inputfolderinfolbl = ('Файлов в папке: {0}, считаю страницы: {1}'.format(len(DivideInputFilesArray), allpagecount))
        DivideInputDirPCountLbl.config(text = inputfolderinfolbl)
        
    inputfolderinfolbl = ('Файлов в папке: {0}, страниц: {1}'.format(len(DivideInputFilesArray), allpagecount))
    DivideInputDirPCountLbl.config(text = inputfolderinfolbl)
    #MainModeBackBtn.configure(state = NORMAL)
    
def DivideOutputDirChoose():
    global DivideOutputDir
    
    DivideOutputDir = filedialog.askdirectory(title='Выберите папку для счет-фактур')
    if DivideOutputDir:
        DivideOutputDirEntry.configure(state = NORMAL)
        DivideOutputDirEntry.delete(0,END)
        DivideOutputDirEntry.insert(0,str(Path(DivideOutputDir).name))
        DivideOutputDirEntry.configure(state = DISABLED)
        print('Divide: ODC: DivideOutputDir :', DivideOutputDir)
        
        DivideOutputDirCheck()
    else:
        print('Divide: ODC: DivideOutputDir not selected')


def DivideOutputDirCheck():
    global DivideInputFile
    global DivideIsInputSel
    global DivideAllPagesCount

    global DivideOutputDir
    global DivideIsOutputSel
    
    isoutdir = os.path.isdir(DivideOutputDir)
    if isoutdir:
        for k in range(5000):
            outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(k)+'.pdf'))
            fileexists = os.path.isfile(outputfile)
            if fileexists:
                DivideIsOutputSel = False
                msgbxlbl = 'В этой папке уже присутствуют документы из выбранного файла! Удалите, переместите эти документы или выберите другую папку'
                messagebox.showerror("", msgbxlbl)
                break
            else:
                DivideIsOutputSel = True

    DivideCheckIfReady()
    

def DivideCheckIfReady():
    global DivideIsInputSel
    global DivideIsOutputSel
    global DivideReadyToGo
    
    print("Divide: CIR: {0}, {1}".format(DivideIsInputSel,DivideIsOutputSel))
    if DivideIsInputSel and DivideIsOutputSel:
        DivideStartDivisionBtn.configure(state = NORMAL)
        DivideReadyToGo = True
    else:
        DivideReadyToGo = False
        DivideStartDivisionBtn.configure(state = DISABLED)




def DivideStartDivision():
    global DivideReadyToGo
    
    DivideOutputDirCheck()
    print(DivideReadyToGo)
    
    if DivideReadyToGo:
        minerthread = Thread(target=DividerProcessManager)
        minerthread.start()
        timethread = Thread(target=DivideTimeUpdater)
        timethread.start()
        minerthread = ""
        timethread = ""


def DividerProcessManager():
    global DivideInputFilesArray
    global DivideOutputDir
    global DivideIsRunning
    global DivideStartedTime
    global MaxProcessCount

    MaxProcessCount = 5 # from zero
    
    DivideBlockGUI(True)
    DivideIsRunning = True
    DivideStartedTime = time.time()
    
    # ProcessManager:
    runningprocessescount = 0
    fileresultarray = []
    dataq = Queue()

    for i in range (len(DivideInputFilesArray)):
        fileresultarray.append(2) # 2 -> awaiting, 1 -> in work, 0 -> ready
    print('ProcessManager: fileresultarray: {0}'.format(fileresultarray))

    processguinumarray = []
    processjustterminated = False
    firstloop = True
    while True:

        # Checking if data received
        if not dataq.empty():
            #### dataq: processnum, status, inputfile, message
            processresult = dataq.get()

            # Refreshing GUI with received data
            for g in range (len(processguinumarray)):
                if processresult[0] == processguinumarray[g]:
                    processguinum = g
                    break
                    
            processresultlbl = ['Документ №{0}: {1}'.format(processresult[0], Path(processresult[2]).name), str(processresult[3])]
            if processguinum == 0:
                DivideProcess1StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 1:
                DivideProcess2StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 2:
                DivideProcess3StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 3:
                DivideProcess4StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 4:
                DivideProcess5StatusLbl.config(text = "\n".join(processresultlbl))
            elif processguinum == 5:
                DivideProcess6StatusLbl.config(text = "\n".join(processresultlbl))
                
                
            if processresult[1] == 0: # process terminated
                processjustterminated = True
                print('ProcessManager: reciever: process {0} terminated'.format(processresult[0]))
                
                runningprocessescount = runningprocessescount - 1
                fileresultarray[processresult[0]] = 0
                
                
                
        # Checking space for new process start
        if processjustterminated or firstloop == True:
            processjustterminated = False
            
            processestostart = MaxProcessCount - runningprocessescount
            
            # Checking amount of files that remains for exec
            filesawaiting = 0
            for r in range (len(fileresultarray)): # Scanning for awaiting files
                if fileresultarray[r] == 2:
                    filesawaiting = filesawaiting + 1
                    
            print('')
            print('')
            print('processjustterminated or firstloop')
            print('ProcessManager: runningprocessescount: {0}'.format(runningprocessescount))
            print('ProcessManager: filesawaiting: {0}'.format(filesawaiting))
            print('ProcessManager: processestostart: {0} ({1})'.format(processestostart, processestostart+1))

            if firstloop:
                firstloop = False
                if filesawaiting <= MaxProcessCount:
                    processestostart = filesawaiting - 1
                else:
                    processestostart = MaxProcessCount
                    
                if filesawaiting == 1:
                    processguinumarray.clear()
                    processguinumarray.append(0)
                    
                needtostartprocess = True
                print('ProcessManager: +++ First loop !')
            elif processestostart <= filesawaiting and  filesawaiting!=0 :
                needtostartprocess = True
                print('ProcessManager: +++ Need to start new process')
            else:
                needtostartprocess = False
            
            if needtostartprocess:
                needtostartprocess = False
                # Iterations for process start
                for s in range (processestostart+1):
                    for j in range (len(fileresultarray)):
                        if fileresultarray[j] == 2: # Scanning for awaiting files
                            fileresultarray[j] = 1 # Mark this file to "1 -> in work"
                            runningprocessescount = runningprocessescount + 1
                            creatingprocessfilenum = j
                            
                            print('ProcessManager: GUI procstart: before processguinumarray {0}'.format(processguinumarray))
                            if runningprocessescount > 1:
                                processguinumarray.clear()
                                for z in range (len(fileresultarray)):
                                    if fileresultarray[z] == 1:
                                        processguinumarray.append(z)
                            print('ProcessManager: GUI procstart: after processguinumarray {0}'.format(processguinumarray))
                            
                            break
                    #### DividerEm: (processnum, dataq, inputfile, outputdir)
                    subprocess = Process(target=DivideMiner, args=(creatingprocessfilenum, dataq, DivideInputFilesArray[creatingprocessfilenum], DivideOutputDir))
                    subprocess.start()
                    print('ProcessManager: ***** Just started new process ! {0}'.format(j))
                    print('ProcessManager: **** runningprocessescount: {0}'.format(runningprocessescount))
                    print('ProcessManager: **** fileresultarray {0}'.format(fileresultarray))
            print('')
            print('')
            
            
            
        # Checking if all is done
        filesready = 0
        for e in range (len(fileresultarray)): # Scanning for ready files
            if fileresultarray[e] == 0:
                filesready = filesready + 1
                
        
        
        if filesready == len(fileresultarray):
            DivideStatusLbl.config(text = 'Завершено! {0} из {1}'.format(filesready, len(fileresultarray)))
            messagebox.showinfo("", "Обработка завершена !")
            print('ProcessManager: All done !!!')
            break
        else:
            DivideStatusLbl.config(text = 'Выполнено {0} из {1}'.format(filesready, len(fileresultarray)))

    DivideIsRunning = False
    DivideBlockGUI(False)


def DivideMiner(processnum, dataq, DivideInputFile, DivideOutputDir):

    processname = current_process().name
    print("= Divider №{0}--{1} STARTED !: {2}".format(processnum, processname, DivideInputFile))
    
    word = 'Счет-фактура №'
    pagecounter = 1
    finded = False
    firstpagesarray = []
    firstpagesarray.clear

    pdf = PdfFileReader(DivideInputFile)
    DivideInputPagesCount = pdf.getNumPages()

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
                        firstpagesarray.append(pagecounter)
                        print("= Divider №{0}--{1}, file: {2}    finded! page {3}".format(processnum, processname, DivideInputFile, pagecounter))
                        finded = True
        if finded:
            finded = False
        else:
            print("= Divider №{0}--{1}, file: {2}    page {3}".format(processnum, processname, DivideInputFile, pagecounter))
            
        progresslbltxt = "Поиск первых страниц... Чтение {0} из {1}, найдено: {2}".format(pagecounter, DivideInputPagesCount, len(firstpagesarray))
        dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)
        pagecounter = pagecounter + 1
        
    print("= Divider №{0}--{1}, ENDED !: {2}".format(processnum, processname, DivideInputFile))
    
    pdftomine = ''
    manager = ''
    laparams = ''
    dev = ''
    interpreter = ''
    pages = ''
    
    
    if len(firstpagesarray) == 0:
        progresslbltxt = "Поиск завершен, счет-фактур не найдено !"
        dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)
        msgbxlbl = 'В выбранном документе не найдено счет-фактур !'
        messagebox.showerror("", msgbxlbl)
    else:
        print("= Divider №{0}--{1}, FM Started !: {2}".format(processnum, processname, DivideInputFile))
        
        originalpdf = PdfFileReader(DivideInputFile)
       
        for k in range(len(firstpagesarray)):
            if k+1 < len(firstpagesarray):
            
                print('**** Документ №: ', str(k+1))
                outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(firstpagesarray[k])+'.pdf'))
                print("Номер первой стр: {0}, номер сл.первой {1}".format(firstpagesarray[k],firstpagesarray[k+1]))
                print("Итоговый файл: {0}".format(outputfile))
                print('Список страниц документа:')
                
                temparray = []
                temparray.clear()
                for x in range(firstpagesarray[k], firstpagesarray[k+1]):
                    print(x)
                    temparray.append(x)
                
                pdf_writer = PdfFileWriter()
                for x in range(len(temparray)):
                    pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

                progresslbltxt = "Создание документов... Обработка {0} из {1}".format(k, len(firstpagesarray))
                dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
                #DivideStatusLbl.config(text = progresslbltxt)

                pdf_writer.write(open(outputfile, 'wb'))
                pdf_writer = ''

                print('Обработка завершена')
                print('*******************')
                print('')


        print('**** Последний документ: ', str(len(firstpagesarray)))
        outputfile = Path (DivideOutputDir, (str(Path(DivideInputFile).name)+" - стр."+str(firstpagesarray[k])+'.pdf'))
        print("Итоговый файл: {0}".format(outputfile))
        print('Список страниц документа:')
        
        temparray.clear()
        for x in range(firstpagesarray[len(firstpagesarray)-1], DivideInputPagesCount+1):
            print(x)
            temparray.append(x)
            
        pdf_writer = PdfFileWriter()
        for x in range(len(temparray)):
            pdf_writer.addPage(originalpdf.getPage(temparray[x]-1))

        progresslbltxt = "Создание документов... Обработка {0} из {1}".format(len(firstpagesarray), len(firstpagesarray))
        dataq.put([processnum, 1, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)

        pdf_writer.write(open(outputfile, 'wb'))
        pdf_writer = ''
        
        progresslbltxt = "Обработка завершена!"
        dataq.put([processnum, 0, DivideInputFile, progresslbltxt])
        #DivideStatusLbl.config(text = progresslbltxt)

        print('Обработка завершена')
        print('****************************')

        print("= Divider №{0}--{1}, FM : {2},    Documents in file: {3}".format(processnum, processname, DivideInputFile, len(firstpagesarray)))
        print("= Divider №{0}--{1}, FM Ended !: {2}".format(processnum, processname, DivideInputFile))
    



def DivideBlockGUI(yes):
    if yes:
        #MainModeBackBtn.configure(state = DISABLED)
        DivideInputDirChooseBtn.configure(state = DISABLED)
        DivideOutputDirBtn.configure(state = DISABLED)
        DivideStartDivisionBtn.configure(state = DISABLED)
    else:
        #MainModeBackBtn.configure(state = NORMAL)
        DivideInputDirChooseBtn.configure(state = NORMAL)
        DivideOutputDirBtn.configure(state = NORMAL)
        DivideStartDivisionBtn.configure(state = NORMAL)


def DivideResizeGUI():
    global scrnwparam
    global scrnhparam
    global DivideGUIisResized
    
    if DivideGUIisResized:
        DivideGUIisResized = not DivideGUIisResized
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x225+{}+{}'.format(scrnw, scrnh))
        DivideResizeGUIBtn.config(text = '➘')
    else:
        DivideGUIisResized = not DivideGUIisResized
        scrnw = (root.winfo_screenwidth()//2) - scrnwparam
        scrnh = (root.winfo_screenheight()//2) - scrnhparam
        root.geometry('375x475+{}+{}'.format(scrnw, scrnh))
        DivideResizeGUIBtn.config(text = '➚')


def DivideTimeUpdater():
    global DivideIsRunning
    global DivideStartedTime

    time.sleep(0.5)
    while DivideIsRunning:
        result = time.time() - DivideStartedTime
        result = datetime2.timedelta(seconds=round(result))
        DivideTimeLbl.config(text = str(result))
        time.sleep(0.01)

#########################










def ScanFolder(folder, extension, filearray):
        filearray.clear()
        for file in os.listdir(folder):
            if file.endswith(extension):
                filearray.append(os.path.join(folder, file))

def PDFCountPages(inputfile):
    pdf = PdfFileReader(inputfile)
    pagecount = pdf.getNumPages()
    pdf = ""
    return pagecount
    
def Similar(a, b):
    return SequenceMatcher(None, a, b).ratio()
    
def ResourcePath(relative_path):    
    try:       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    main()

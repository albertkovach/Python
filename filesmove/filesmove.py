from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os
global root
global inputdir
global inputdirtxt
global savedir
global savedirtxt
global barcodefile
global barcodetxt

global isinpsel
global isoutsel
global isbarsel

global FilesArray
global ValidFilesCount

global BarcodeArray
global BarcodeCount




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
    root.geometry('375x310+{}+{}'.format(scrnw, scrnh))
        
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
        global isinpsel
        global isoutsel
        global isbarsel
        isinpsel = False
        isoutsel = False
        isbarsel = False
        
    
        global InputDirPathLbl
        InputDirPathLbl = Label(text="Выберите папку с файлами:", background="white", font=("Arial", 10))
        InputDirPathLbl.place(x=16, y=7)
        
        global InputDirEntry
        global inputdirtxt
        inputdirtxt = StringVar()
        InputDirEntry = Entry(fg="black", bg="white", width=46, textvariable=inputdirtxt)
        InputDirEntry.place(x=20, y=30)
        InputDirEntry.configure(state = DISABLED)
        
        global InputDirChooseBtn
        InputDirChooseBtn = Button(text='Выбор', command=ChooseInputDir)
        InputDirChooseBtn.place(x=305, y=29, height=20)

        global FilesCountLbl
        FilesCountLbl = Label(text="", background="white")
        FilesCountLbl.place(x=17, y=48)
        
        
        
        global SaveDirBtnLbl
        SaveDirBtnLbl = Label(text="Папка для коробов:", background="white", font=("Arial", 10))
        SaveDirBtnLbl.place(x=16, y=87)
        
        global SaveDirBtn
        SaveDirBtn = Button(text="Выбор", command=ChooseOutputDir)
        SaveDirBtn.place(x=20, y=110, height=20)

        global SaveDirEntry
        global savedirtxt
        savedirtxt = StringVar()
        SaveDirEntry = Entry(fg="black", bg="white", width=46, textvariable=savedirtxt)
        SaveDirEntry.place(x=75, y=111)
        SaveDirEntry.configure(state = DISABLED)
        
        
        
        global BarcodeSelLbl
        BarcodeSelLbl = Label(text="Файл с штрихкодами:", background="white", font=("Arial", 10))
        BarcodeSelLbl.place(x=16, y=141)
        
        global BarcodeSelBtn
        BarcodeSelBtn = Button(text="Выбор", command=ChooseBarcodeFile)
        BarcodeSelBtn.place(x=20, y=164, height=20)

        global BarcodeSelEntry
        global barcodetxt
        barcodetxt = StringVar()
        BarcodeSelEntry = Entry(fg="black", bg="white", width=46, textvariable=barcodetxt)
        BarcodeSelEntry.place(x=75, y=165)
        BarcodeSelEntry.configure(state = DISABLED)
        
        global BarcodeCountLbl
        BarcodeCountLbl = Label(text="", background="white")
        BarcodeCountLbl.place(x=17, y=185)
        
        
        
        global RunBtn
        RunBtn = Button(text="Выполнить", command=StartMoving)
        RunBtn.place(x=20, y=225, width=90, height=25)
        

def ChooseInputDir():
    global inputdir
    inputdir = filedialog.askdirectory()
    
    if inputdir:
        inputdirtxt.set(str(inputdir))
        print('inputdir :', inputdir)
        
        global FilesArray
        global ValidFilesCount
        
        FilesArray = []
        FilesArray.clear()
        
        for file in os.listdir(inputdir):
            if file.endswith(".pdf"):
                FilesArray.append(os.path.join(inputdir, file))
                #print('file:', str(file))
        ValidFilesCount = len(FilesArray)
        FilesCountLbl.config(text = 'Количество файлов PDF: ' + str(len(FilesArray)))
        print('number of valid files -', str(ValidFilesCount))
        
        isinpsel = True
    else:
        print('inputdir is empty')
        




   
def ChooseOutputDir():
    global savedir
    global savedirtxt
    savedir = filedialog.askdirectory()

    if savedir:
        savedirtxt.set(str(savedir))
        print('savedir :', savedir)
        
        isoutsel = True
    else:
        print('savedir is empty')




def ChooseBarcodeFile():
    global barcodefile
    global barcodetxt
    global BarcodeArray
    
    barcodefile = filedialog.askopenfilename(filetypes=(('text files', 'txt'),))
    
    if barcodefile:
        barcodetxt.set(str(barcodefile))
        print('barcodetxt :', barcodefile)
        
        file = open(barcodefile,'r')
        BarcodeArray = []
        BarcodeArray.clear()
        for line in file:
            BarcodeArray.append(line.strip()) # We don't want newlines in our list, do we?
            
        for i in range(len(BarcodeArray)):
            print('bar:', str(BarcodeArray[i]))
            
        BarcodeCountLbl.config(text = 'Кодов в файле: ' + str(len(BarcodeArray)))
        
        isbarsel = True
        
    else:
        print('barcodefile is empty')
    

   
def StartMoving():
    print('StartMoving')
    

if __name__ == '__main__':
    main()

from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import os, sys, re, shutil
from pathlib import Path

from wand.image import Image
from wand.display import display
from wand.drawing import Drawing

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

def main():
    global root
    global bckgcolor
    bckgcolor = '#F5F5F5'

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('590x265+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background=bckgcolor)   
        self.parent = parent
        self.parent.title("Wand")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global AddLogo
        global AddText
        global AddTextFile
        
        AddLogo = IntVar()
        AddLogo.set(0)
        AddText = IntVar()
        AddText.set(0)
        AddTextFile = IntVar()
        AddTextFile.set(0)
        
        global InputFile
        global LogoFile
        global OutputFolder
        
        InputFile = ""
        LogoFile = ""
        OutputFolder = ""
        
    
        global MainLbl
        MainLbl = Label(text="Выбор основного файла:", bg=bckgcolor, font=("Arial", 13))
        MainLbl.place(x=16, y=10)
        
        global InputFileEntry
        InputFileEntry = Entry(fg="black", bg=bckgcolor, width=21, font=("Arial", 11))
        InputFileEntry.place(x=20, y=37)
        
        global InputFileBtn
        InputFileBtn = Button(text='...', bg=bckgcolor, command=SelectFile, font=("Arial", 9))
        InputFileBtn.place(x=200, y=36, width=22, height=22)
        
        
        
        
        global OutputLbl
        OutputLbl = Label(text="Папка для итоговых файлов:", bg=bckgcolor, font=("Arial", 12))
        OutputLbl.place(x=36, y=71)
        
        global OutputDirEntry
        OutputDirEntry = Entry(fg="black", bg=bckgcolor, width=21, font=("Arial", 11))
        OutputDirEntry.place(x=40, y=97)
        
        global OutputDirBtn
        OutputDirBtn = Button(text='...', bg=bckgcolor, command=SelectFile, font=("Arial", 10))
        OutputDirBtn.place(x=220, y=96, width=22, height=22)
        
        
        
        
        global ChBxAddLogo
        ChBxAddLogo = Checkbutton(text="Добавить логотип", bg=bckgcolor, variable=AddLogo, onvalue=1, offvalue=0, font=("Arial", 11), command=InterfaceTrigger)
        ChBxAddLogo.place(x=35, y=135)
        
        global LogoFileEntry
        LogoFileEntry = Entry(fg="black", bg=bckgcolor, width=20, font=("Arial", 11))
        LogoFileEntry.place(x=60, y=165)
        
        global LogoFileBtn
        LogoFileBtn = Button(text='...', command=SelectLogoFile, font=("Arial", 10))
        LogoFileBtn.place(x=230, y=163, width=22, height=22)
        
        global LogoParamLbl
        LogoParamLbl = Label(text="Параметры: X, Y, Масштаб", bg=bckgcolor, font=("Arial", 10))
        LogoParamLbl.place(x=57, y=190)
        
        global LogoXEntry
        LogoXEntry = Entry(fg="black", bg=bckgcolor, width=5, font=("Arial", 11))
        LogoXEntry.place(x=60, y=215)
        LogoXEntry.insert(0, "50")
        
        global LogoYEntry
        LogoYEntry = Entry(fg="black", bg=bckgcolor, width=5, font=("Arial", 11))
        LogoYEntry.place(x=120, y=215)
        LogoYEntry.insert(0, "50")
        
        global LogoScaleEntry
        LogoScaleEntry = Entry(fg="black", bg=bckgcolor, width=5, font=("Arial", 11))
        LogoScaleEntry.place(x=180, y=215)
        LogoScaleEntry.insert(0, "1")
        
        
        
        
        global ChBxAddText
        ChBxAddText = Checkbutton(text="Добавить текст", bg=bckgcolor, variable=AddText, onvalue=1, offvalue=0, font=("Arial", 11), command=InterfaceTrigger)
        ChBxAddText.place(x=305, y=40)
        
        global ChBxAddTextFile
        ChBxAddTextFile = Checkbutton(text="Выбрать файл с текстом", bg=bckgcolor, variable=AddTextFile, onvalue=1, offvalue=0, font=("Arial", 10), command=InterfaceTrigger)
        ChBxAddTextFile.place(x=326, y=64)
        
        global TextFileEntry
        TextFileEntry = Entry(fg="black", bg=bckgcolor, width=20, font=("Arial", 11))
        TextFileEntry.place(x=330, y=94)
        
        global TextFileBtn
        TextFileBtn = Button(text='...', command=SelectLogoFile, font=("Arial", 10))
        TextFileBtn.place(x=500, y=93, width=22, height=22)
        
        
        global ChBxAddNameFile
        ChBxAddNameFile = Checkbutton(text="Выбрать файл с именами", bg=bckgcolor, variable=AddTextFile, onvalue=1, offvalue=0, font=("Arial", 10), command=InterfaceTrigger)
        ChBxAddNameFile.place(x=326, y=116)
        
        global NameFileEntry
        NameFileEntry = Entry(fg="black", bg=bckgcolor, width=20, font=("Arial", 11))
        NameFileEntry.place(x=330, y=146)
        
        global NameFileBtn
        NameFileBtn = Button(text='...', command=SelectLogoFile, font=("Arial", 10))
        NameFileBtn.place(x=500, y=145, width=22, height=22)
        
        
        global TextParamLbl
        TextParamLbl = Label(text="Параметры: X, Y, Размер", bg=bckgcolor, font=("Arial", 10))
        TextParamLbl.place(x=327, y=117)
        
        global TextXEntry
        TextXEntry = Entry(fg="black", bg=bckgcolor, width=5, font=("Arial", 11))
        TextXEntry.place(x=330, y=142)
        TextXEntry.insert(0, "100")
        
        global TextYEntry
        TextYEntry = Entry(fg="black", bg=bckgcolor, width=5, font=("Arial", 11))
        TextYEntry.place(x=390, y=142)
        TextYEntry.insert(0, "100")
        
        global TextScaleEntry
        TextScaleEntry = Entry(fg="black", bg=bckgcolor, width=5, font=("Arial", 11))
        TextScaleEntry.place(x=450, y=142)
        TextScaleEntry.insert(0, "40")
        
        
        
        
        global RunBtn
        RunBtn = Button(text='Запуск', command=Run, font=("Arial", 12))
        RunBtn.place(x=300, y=190, width=90, height=32)
        
        global PreviewBtn
        PreviewBtn = Button(text='Предпросмотр', command=OpenPreview, font=("Arial", 12))
        PreviewBtn.place(x=410, y=190, width=140, height=32)
        
        
        InterfaceTrigger()





def SelectFile():
    global InputFile

    InputFile = ""
    InputFile = filedialog.askopenfilename(filetypes=[("Excel files", ".png")])
    if InputFile:
        InputFileEntry.configure(state = NORMAL)
        InputFileEntry.delete(0,END)
        InputFileEntry.insert(0,str(Path(InputFile).name))
        InputFileEntry.configure(state = DISABLED)
        print('IDC: InputFile : {0}'.format(InputFile))
    else:
        print('IDC: InputFile not selected')


def SelectLogoFile():
    global LogoFile

    LogoFile = ""
    LogoFile = filedialog.askopenfilename(filetypes=[("Excel files", ".png")])
    if LogoFile:
        LogoFileEntry.configure(state = NORMAL)
        LogoFileEntry.delete(0,END)
        LogoFileEntry.insert(0,str(Path(LogoFile).name))
        LogoFileEntry.configure(state = DISABLED)
        print('IDC: LogoFile : {0}'.format(LogoFile))
    else:
        print('IDC: LogoFile not selected')


def SelectOutput():
    global OutputFolder

    OutputFolder = ""
    OutputFolder = filedialog.askdirectory(title='Выберите папку на обработку')
    if OutputFolder:
        OutputDirEntry.configure(state = NORMAL)
        OutputDirEntry.delete(0,END)
        OutputDirEntry.insert(0,str(OutputFolder))
        OutputDirEntry.configure(state = DISABLED)
        print('IDC: OutputFolder : {0}'.format(OutputFolder))
    else:
        print('IDC: OutputFolder not selected')


def OpenPreview():
    global InputFile
    
    outputfolderpath = Path(Path(InputFile).parent, "out")
    isoutputfolderpath = os.path.isdir(outputfolderpath)
    if not isoutputfolderpath:
        os.makedirs(outputfolderpath)
    RunWand(outputfolderpath)
    
    os.system('"{0}"'.format(Path(outputfolderpath, Path(InputFile).name)))



def Run():
    global OutputFolder
    global InputFile
    
    isoutputfolder = os.path.isdir(OutputFolder)
    if isoutputfolder:
        RunWand(OutputFolder)
    else:
        defaultfolderpath = Path(Path(InputFile).parent, "out")
        isdeffolderpath = os.path.isdir(defaultfolderpath)
        if not isdeffolderpath:
            os.makedirs(defaultfolderpath)
        RunWand(defaultfolderpath)
    
    

def RunWand(output):
    global InputFile
    global LogoFile
    
    with Image(filename=InputFile) as image:
        if AddLogo.get() == 1:
            with Image(filename=LogoFile) as fg_img:
                tempfolderpath = Path(Path(InputFile).parent, "temp")
                istempfolderpath = os.path.isdir(tempfolderpath)
                fg_img.resize(int(fg_img.width*float(LogoScaleEntry.get())), int(fg_img.height*float(LogoScaleEntry.get())))
                image.composite(fg_img, left=int(LogoXEntry.get()), top=int(LogoYEntry.get()))
        if AddText.get() == 1:
            with Drawing() as draw:
                draw.font = 'wandtests/assets/League_Gothic.otf'
                #if AddTextFile.get() == 0:
                    #with open(r"myfile.txt", 'r') as fp:
                    #lines = len(fp.readlines())
                    #print('Total Number of lines:', lines)
                text = TextFileEntry.get()
                draw.font_size = int(TextScaleEntry.get())
                draw.text(int(TextXEntry.get()), int(TextYEntry.get()), text)
                draw(image)
        image.format = 'png'
        image.save(filename=Path(output, Path(InputFile).name))




def InterfaceTrigger():

    if AddLogo.get() == 1:
        LogoFileEntry.configure(state = NORMAL)
        LogoFileBtn.configure(state = NORMAL)
        LogoParamLbl.configure(state = NORMAL)
        LogoXEntry.configure(state = NORMAL)
        LogoYEntry.configure(state = NORMAL)
        LogoScaleEntry.configure(state = NORMAL)
    else:
        LogoFileEntry.configure(state = DISABLED)
        LogoFileBtn.configure(state = DISABLED)
        LogoParamLbl.configure(state = DISABLED)
        LogoXEntry.configure(state = DISABLED)
        LogoYEntry.configure(state = DISABLED)
        LogoScaleEntry.configure(state = DISABLED)


    if AddText.get() == 1:
        ChBxAddTextFile.configure(state = NORMAL)
        TextFileEntry.configure(state = NORMAL)
        TextFileBtn.configure(state = NORMAL)
        TextParamLbl.configure(state = NORMAL)
        TextXEntry.configure(state = NORMAL)
        TextYEntry.configure(state = NORMAL)
        TextScaleEntry.configure(state = NORMAL)
    else:
        ChBxAddTextFile.configure(state = DISABLED)
        TextFileEntry.configure(state = DISABLED)
        TextFileBtn.configure(state = DISABLED)
        TextParamLbl.configure(state = DISABLED)
        TextXEntry.configure(state = DISABLED)
        TextYEntry.configure(state = DISABLED)
        TextScaleEntry.configure(state = DISABLED)
        
        
    if AddTextFile.get() == 1:
        TextFileBtn.place(x=500, y=93, width=22, height=22)
        
        ChBxAddNameFile.place(x=326, y=116)
        NameFileEntry.place(x=330, y=146)
        NameFileBtn.place(x=500, y=145, width=22, height=22)
        
        TextParamLbl.place(x=327, y=169)
        TextXEntry.place(x=330, y=194)
        TextYEntry.place(x=390, y=194)
        TextScaleEntry.place(x=450, y=194)
        RunBtn.place(x=300, y=242-20, width=90, height=32)
        PreviewBtn.place(x=410, y=242-20, width=140, height=32)
    else:
        TextFileBtn.place_forget()
        
        ChBxAddNameFile.place_forget()
        NameFileEntry.place_forget()
        NameFileBtn.place_forget()
        
        TextParamLbl.place(x=327, y=117)
        TextXEntry.place(x=330, y=142)
        TextYEntry.place(x=390, y=142)
        TextScaleEntry.place(x=450, y=142)
        RunBtn.place(x=300, y=190, width=90, height=32)
        PreviewBtn.place(x=410, y=190, width=140, height=32)



if __name__ == '__main__':
    main()

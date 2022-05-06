from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import win32api
import win32print
import PyPDF2
from pathlib import Path
import os, sys

global root
global scrnwparam
global scrnhparam
scrnwparam = 125
scrnhparam = 110




def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    datafile = "icon.ico"
    if not hasattr(sys, "frozen"):
        datafile = os.path.join(os.path.dirname(__file__), datafile)
    else:
        datafile = os.path.join(sys.prefix, datafile)
    root.iconbitmap(default=resource_path(datafile))
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('250x80+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("Печать папки")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global TextLbl
        TextLbl = Label(text="", background="white")
        TextLbl.place(x=17, y=50)
        
        global ChooseFileBtn
        ChooseFileBtn = Button(text='Выбрать папку с PDF', command=BtnCmd1)
        ChooseFileBtn.place(x=18, y=20, height=25)

        global RunBtn
        RunBtn = Button(text='Печать', command=BtnCmd2)
        RunBtn.place(x=180, y=20, height=25)
        RunBtn.configure(state = DISABLED)


def BtnCmd1():
    global directory
    directory = filedialog.askdirectory()
    
    if directory:
        print('Selected:', directory)
        
        global FilesArray
        global ValidFilesCount
        
        FilesArray = []
        FilesArray.clear()
        
        filenum = 0
        for file in os.listdir(directory):
            if file.endswith(".pdf"):
                filenum = filenum + 1
                FilesArray.append(os.path.join(directory, file))
                print('{0}: {1}'.format(filenum, file))
        ValidFilesCount = len(FilesArray)
        if ValidFilesCount > 0:
            TextLbl.config(text = 'Файлов PDF: {0}, папка: {1}'.format(len(FilesArray), Path(directory).name))
            print('Количество подходящих файлов - {0}'.format(ValidFilesCount))
            RunBtn.configure(state = NORMAL)
        else:
            TextLbl.config(text = 'В папке нет файлов PDF !')
            print('Количество подходящих файлов - {0}'.format(ValidFilesCount))


def BtnCmd2():
    global filename
    global FilesArray
    global ValidFilesCount
    
    
    answer = messagebox.askyesno(
        title="Подтверждение", 
        message="Вы уверены ?")
    if answer:

        try:
            for i in range(ValidFilesCount):
                win32api.ShellExecute(0, "print", FilesArray[i], None,  ".",  0)
                print('Отправлен на печать: {0}'.format(FilesArray[i]))
                
            RunBtn.configure(state = DISABLED)
            msgbxlbl = 'Файлы отправлены на печать !'
            messagebox.showinfo("Выполнено !", msgbxlbl)
            print('Выполнено !')
        except:
            print("Ошибка печати !")
            msgbxlbl = 'Ошибка печати !'
            messagebox.showerror("Ошибка !", msgbxlbl)


    
    
def resource_path(relative_path):    
    try:       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
    

if __name__ == '__main__':
    main()

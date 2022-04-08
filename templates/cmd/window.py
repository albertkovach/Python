from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from threading import Thread

import subprocess
import os

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
    root.geometry('170x130+{}+{}'.format(scrnw, scrnh))

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
        #global SLbl
        #Lbl = Label(text="Label", background="white")
        #Lbl.place(x=16, y=10)
        
        #global SEntry
        #SEntry = Entry(fg="black", bg="white", width=20)
        #SEntry.place(x=20, y=32)
        
        global SBtn
        SBtn = Button(text='Button', command=BtnCmd)
        SBtn.place(x=20, y=20, height=20)

def BtnCmd():
    print('btn pressed')
    threada = Thread(target=StartSub)
    threadb = Thread(target=StartSub)
    threadc = Thread(target=StartSub)
    threada.start()
    threadb.start()
    threadc.start()
    #os.system("cmdpdfminer.py 1")




def StartSub():
    os.system('cmd.exe /c "C:\\Users\\ORB User\\Desktop\\cmdpdfminer2.exe"')
    #os.system('cmd.exe')
    
    #subprocess.Popen("cmdpdfminer.py 1", shell=True)













if __name__ == '__main__':
    main()

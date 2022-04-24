from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root, canv
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

import time


def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('400x400+{}+{}'.format(scrnw, scrnh))
        
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
        global root, canv
        
        canv = Canvas(self, bg='white')
        coord = 10, 10, 300, 300
        
        innerdia = 230
        widthdiff = (coord[2]-coord[0]-innerdia)/2
        heigthdiff = (coord[3]-coord[1]-innerdia)/2
        whitefill = widthdiff+coord[0], heigthdiff+coord[1], coord[2]-widthdiff, coord[3]-heigthdiff
        
        global arc
        arc = canv.create_arc(coord, start=0, extent=200, fill="green", outline='white')
        canv.create_oval(whitefill, fill='white', outline='white')
        canv.pack(fill=BOTH, expand=1)
        
        canv.itemconfig(arc, extent=100)
        
        ArcAnim()
        
        

def ArcAnim():
    global canv
    ns = 0
    
    while True:
        ns = ns + 1
        time.sleep(1)
        canv.itemconfig(arc, extent=ns)

if __name__ == '__main__':
    main()

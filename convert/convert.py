from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

from pathlib import Path
from pathlib import PurePath

import pandas as pd


import tabula
import pdftables_api

global root
global scrnwparam
global scrnhparam
scrnwparam = 138
scrnhparam = 110

def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('275x110+{}+{}'.format(scrnw, scrnh))
        
    app = GUI(root)
    root.mainloop()

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background="white")   
        self.parent = parent
        self.parent.title("Resolver")
        self.pack(fill=BOTH, expand=1)
        self.initUI()
    
    def initUI(self):
        global InputFileLbl
        InputFileLbl = Label(text="Выберите УПД:", background="white", font=("Arial", 10))
        InputFileLbl.place(x=16, y=10)
        
        global InputFileEntry
        InputFileEntry = Entry(fg="black", bg="white", width=30)
        InputFileEntry.place(x=20, y=32)
        InputFileEntry.configure(state = DISABLED)
        
        global InputFileBtn
        InputFileBtn = Button(text='Выбор', command=InputFileChoose)
        InputFileBtn.place(x=210, y=31, height=20)
        
        global StartBtn
        StartBtn = Button(text='Магия !', command=Start)
        StartBtn.place(x=20, y=65, width=75, height=25)
        StartBtn.configure(state = DISABLED)



def InputFileChoose():
    global InputFilePath
    global OutputFilePath
    global IsInputSel

    InputFilePath = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
    if InputFilePath:
        InputFileEntry.configure(state = NORMAL)
        InputFileEntry.delete(0,END)
        InputFileEntry.insert(0,str(Path(InputFilePath).name))
        InputFileEntry.configure(state = DISABLED)
        print('IFC: InputFilePath :', InputFilePath)
        
        #OutputFilePath = Path(Path(InputFilePath).parent, Path(InputFilePath).stem, '.xlsx').as_posix()
        #OutputFilePath = Path(Path(InputFilePath).parent, '111.csv').as_posix()
        OutputFilePath = Path(Path(InputFilePath).parent, 'XLS {0}.csv'.format(Path(InputFilePath).stem)).as_posix()
        
        
        print('IFC: OutputFilePath :', OutputFilePath)
        
        IsInputSel = True
        StartBtn.configure(state = NORMAL)
    else:
        print('IFC: InputFilePath not selected')


def Start():
    global InputFilePath
    global OutputFilePath
    
    global IsInputSel
    
    
    #df = tabula.read_pdf(InputFilePath, pages="all")
    
    #tabula.convert_into(InputFilePath, OutputFilePath, output_format="csv")
    
    #df = tabula.read_pdf(InputFilePath, pages = 1)[0]
    #df.to_excel(OutputFilePath)
    

    #PDF = tabula.read_pdf(InputFilePath, pages='all', multiple_tables=True)
    #PDF = pd.DataFrame(PDF)
    #PDF.to_excel(OutputFilePath, index=False) 
    #PDF.to_csv(OutputFilePath, sep='\t', encoding='utf-8')
    
    
    dfs = tabula.read_pdf(InputFilePath, pages='all')
    tabula.convert_into(InputFilePath, OutputFilePath, output_format="csv", pages='all')
    


if __name__ == '__main__':
    main()

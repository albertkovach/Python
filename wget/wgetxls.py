from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

import requests, os, sys, csv
from pathlib import Path

import xml.etree.ElementTree as ET

from openpyxl import Workbook

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150


def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('240x130+{}+{}'.format(scrnw, scrnh))
        
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
    
        global Debug
        Debug = True
    
        global SLbl
        Lbl = Label(text="Папка проекта:", background="white")
        Lbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=20)
        InputDirEntry.place(x=20, y=32)
        
        global InputDirBtn
        InputDirBtn = Button(text='Выбор', command=OpenXml)
        #InputDirBtn.place(x=20, y=54, height=20)
        InputDirBtn.place(x=150, y=31, height=20)

        global RunBtn
        RunBtn = Button(text='Run', command=PrintAllXml)
        RunBtn.place(x=20, y=60, height=20)


def BtnCmd():
    price_url = "https://shop.feron.ru/bitrix/catalog_export/im.xml"
    
    req = requests.get(price_url)
    reqfilepath = Path('C:\\Users\\user\\Documents\\GitHub\\Python\\wget\\feron.xml')
    revisedreqfile = reqfilepath.as_posix()
    
    with open(revisedreqfile,'wb') as reqfile:
        reqfile.write(req.content)
    
    
    
def OpenXml():
    #reqfile = Path('C:\\Users\\user\\Documents\\GitHub\\Python\\wget\\R60.xml')
    reqfilepath = Path('C:\\Users\\user\\Documents\\GitHub\\Python\\wget\\feron.xml')
    revisedreqfile = reqfilepath.as_posix()
    
    tree = ET.parse(revisedreqfile)
    root = tree.getroot()
    
    wbpath = Path('C:\\Users\\user\\Documents\\GitHub\\Python\\wget\\table.xlsx')
    wb = Workbook()
    wscategory = wb.create_sheet("Category")
    wsitems = wb.create_sheet("Items")
    

    categoryfieldnames = ['category', 'id', 'parentId']
    wscategory.append(categoryfieldnames)
    
    itemsfieldnames = ['description', 'vendorCode', 'vendor', 'picture', 'categoryId', 'url', 'price mitc']
    wsitems.append(itemsfieldnames)
    
    
    for L1 in root:
        for L2 in L1:
            for L3 in L2:
                
                if L3.tag == 'category':
                    catname = L3.text
                    catid = L3.get('id')
                    catparentid = L3.get('parentId')
                    if Debug: print("|-> Category: ", catid, catparentid, catname)
                    wscategory.append([catname, catid, catparentid])
                    
                if L3.tag == 'offer':
                    if Debug: print("")
                    if Debug: print("**** Item : ", L3.attrib)
                    
                    for L4 in L3:
                        if L4.tag == 'url':
                            if Debug: print("====== url: ", L4.text)
                            itemrow = []
                            itempics = []
                            itemrow.append(L4.text)
                        if L4.tag == 'categoryId':
                            if Debug: print("====== categoryId: ", L4.text)
                            itemrow.append(L4.text)
                        if L4.tag == 'picture':
                            if Debug: print("====== picture: ", L4.text)
                            itempics.append(L4.text)
                        if L4.tag == 'vendor':
                            if Debug: print("====== vendor: ", L4.text)
                            picstr = ""
                            for pic in itempics:
                                picstr = picstr + pic + "||"
                            itemrow.append(picstr)
                            #itemrow.append('pics')
                            itemrow.append(L4.text)
                        if L4.tag == 'vendorCode':
                            if Debug: print("====== vendorCode: ", L4.text)
                            itemrow.append(L4.text)
                        if L4.tag == 'description':
                            if Debug: print("====== description: ", L4.text)
                            itemrow.append(L4.text)
                        if L4.tag == 'param':
                            name = L4.get('name')
                            if name == 'МИЦ (мин. интернет-цена)':
                                if Debug: print("====== price mitc: ", L4.text)
                                itemrow.append(L4.text)
                                
                                try:
                                    wsitems.append([itemrow[5], itemrow[4], itemrow[3], itemrow[2], itemrow[1], itemrow[0], itemrow[6]])
                                except:
                                    wsitems.append([itemrow[5], itemrow[4], itemrow[3], itemrow[2], itemrow[1], itemrow[0]])
                                    pass
    wb.save(wbpath)

def PrintAllXml():
    reqfilepath = Path('C:\\Users\\user\\Documents\\GitHub\\Python\\wget\\feron.xml')
    revisedreqfile = reqfilepath.as_posix()
    
    tree = ET.parse(revisedreqfile)
    root = tree.getroot()
    
    for L1 in root:
        print("|--> L1: ", L1.tag, L1.attrib, L1.text)
        
        for L2 in L1:
            print("|-----> L2: ", L2.tag, L2.attrib, L2.text)
            
            for L3 in L2:
                print("|--------> L3: ", L3.tag, L3.attrib, L3.text)
                
                for L4 in L3:
                    print("|---------> L4: ", L4.tag, L4.attrib, L4.text)

def SelectDir():
    global InputDir
    
    print('IDC: InputFile not selected')


if __name__ == '__main__':
    main()

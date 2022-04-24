from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

global root
global scrnwparam
global scrnhparam
scrnwparam = 185
scrnhparam = 150

global filename

from pdfminer.layout import LAParams, LTTextBox, LTText, LTChar, LTAnno
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import PDFPageAggregator


def main():
    global root

    root = Tk()
    root.resizable(False, False)
        
    scrnw = (root.winfo_screenwidth()//2) - scrnwparam
    scrnh = (root.winfo_screenheight()//2) - scrnhparam
    root.geometry('300x130+{}+{}'.format(scrnw, scrnh))
        
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
        global InputDirLbl
        InputDirLbl = Label(text="Выберите папку с файлами:", background="white")
        InputDirLbl.place(x=16, y=10)
        
        global InputDirEntry
        InputDirEntry = Entry(fg="black", bg="white", width=20)
        InputDirEntry.place(x=20, y=32, width=180)
        
        global InputDirBtn
        InputDirBtn = Button(text='Выбор', command=InputFileChoose)
        InputDirBtn.place(x=210, y=30, height=20)
        
        
        global ExecuteBtn
        ExecuteBtn = Button(text='Запуск', command=Searcher)
        ExecuteBtn.place(x=20, y=60, width=60, height=20)





def InputFileChoose():
    global filename
    filename = filedialog.askopenfilename(filetypes=(('', 'pdf'),))
    InputDirEntry.config(text = str(filename))
    print('Selected:', filename)


def Searcher():
    ExtractText()


def ExtractText():
    global filename
    
    word = '3905069220'

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
                text = textbox.get_text().replace('\n', '---')
                print("_________________________")
                print("coord: {0}, {1}".format(textbox.bbox[0], textbox.bbox[1]))
                print("text: {0}".format(text))
                print("length: {0}".format(len(text)))
                #for line in textbox:
                #    text = line.get_text()
                #    print(text)
                #    for char in line:
                #        print(char.get_text())


    
    pdftomine = ''
    manager = ''
    laparams = ''
    dev = ''
    interpreter = ''
    pages = ''


def extract_text_from_pdf(pdf_path):
    """
    This function extracts text from pdf file and return text as string.
    :param pdf_path: path to pdf file.
    :return: text string containing text of pdf.
    """
    resource_manager = PDFResourceManager()
    fake_file_handle = io.StringIO()
    converter = TextConverter(resource_manager, fake_file_handle)
    page_interpreter = PDFPageInterpreter(resource_manager, converter)

    with open(pdf_path, 'rb') as fh:
        for page in PDFPage.get_pages(fh, caching=True, check_extractable=True):
            page_interpreter.process_page(page)

        text = fake_file_handle.getvalue()

    # close open handles
    converter.close()
    fake_file_handle.close()

    if text:
        return text
    return None


if __name__ == '__main__':
    main()

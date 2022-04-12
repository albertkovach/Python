import win32api
#import win32print
from win32 import win32print
from glob import glob

# A List containing the system printers
all_printers = [printer[2] for printer in win32print.EnumPrinters(2)]
# Ask the user to select a printer
printer_num = int(input("Choose a printer:\n"+"\n".join([f"{n} {p}" for n, p in enumerate(all_printers)])+"\n"))
# set the default printer
win32print.SetDefaultPrinter(all_printers[printer_num])
pdf_dir = "D:/path/to/pdf_dir/**/*"
for f in glob(pdf_dir, recursive=True):
    win32api.ShellExecute(0, "print", f, None,  ".",  0)

input("press any key to exit")
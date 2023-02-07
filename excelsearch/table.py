from tkintertable import TableCanvas, TableModel
from tkinter import *

root=Tk()
t_frame=Frame(root)
t_frame.pack(fill='both', expand=True)

table = TableCanvas(t_frame, cellwidth=60, thefont=('Arial',12),rowheight=18, rowheaderwidth=30,
                    rowselectedcolor='yellow', editable=True)

table.show()

root.mainloop()
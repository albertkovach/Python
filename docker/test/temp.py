Обработка короба 2600044444411: 8 из 34
	Сортировка файлов: 2 из 70
	Сборка одностраничных документов
	Сборка многостраничных документов
	Формирование отчета
	Удаление временных файлов
    

	
SortInputDirPathLbl.place_forget()
SortInputDirEntry.place_forget()
SortInputDirChooseBtn.place_forget()
SortValidDirsCountLbl.place_forget()
SortStartCombineBtn.place_forget()
SortStatus1Lbl.place_forget()
SortStatus2Lbl.place_forget()
SortTimeLbl.place_forget()

global SortInputDir
global SortInputDirsArray
global SortInputDirsCount
global SortIsInputSel

global SortIsRunning
global SortStartedTime

    global DivideIsRunning
    global DivideStartedTime
    
    DivideIsRunning = True
    DivideStartedTime = time.time()
    
    
DivideCurrentFileLbl
        
        
    msgbxlbl = 'Обработка файла завершена!'
    messagebox.showinfo("", msgbxlbl)
import os
import sys

import ctypes as c
import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter as xl


#To build use
#pyinstaller --add-binary Random.dll;. --add-binary Random32.dll;. --add-data icon.ico;. --windowed --onefile --noconsole --icon=icon.ico main.py
#and uncomment line below
#icon = os.path.join(sys._MEIPASS,"icon.ico")

#importing .dll
if c.sizeof(c.c_voidp) == 4:
    so_file = ".\Random32.dll"
    print("32x")
else:
    so_file = ".\Random.dll"
    print("64x")

my_functions = c.CDLL(so_file)

GetList = my_functions.GetList
GetList.argtypes = (c.c_int,)
GetList.restype = c.py_object

SetSeed = my_functions.SetSeed

CELL_WIDTH = 17
WIN_SIZE = "350x260"

T_SLOW = 0.250
T_FAST = 0.125
#importing .dll

def browseDirectories(entry):
    """Selecting directory and inserting to entry"""
    filename = tk.filedialog.askdirectory(initialdir = "/",
                             title = "Select a directiry to save",
                             )
    entry.delete(0,tk.END)
    entry.insert(0,filename)

def GenerateXLS():
    """Creating .xls spreadshit"""

    header = ("Вероятность P(ai)", "Опыт №1", "Опыт №2", "Опыт №3", "Опыт №4", "Опыт №5", "Усреднение по 3", "Усреднение по 5")
    alphabet = ("A", "B", "C", "D", "E", "F", "G", "H", "I",
                "J", "K", "L", "M", "N", "O", "P", "Q", "R",
                "S", "T", "U", "V", "W", "X", "Y", "Z",
                "1", "2", "3", "4", "5", "6", "7", "8", "9", "0")
    #Getting entropy
    H = (e_H1.get(),e_H2.get(),e_H3.get())
    # Getting data and handling errors
    try:
        Kn = int(e_Kn.get())
    except:
        tk.messagebox.showerror("Ошибка", "Kn должно быть целым числом до от 1 до 36")
        return
    if (Kn > 36) or (Kn < 1):
        tk.messagebox.showerror("Ошибка", "Kn должно быть целым числом до от 1 до 36")
        return

    try:
        n = int(e_n.get())
    except:
        tk.messagebox.showerror("Ошибка", "n должно быть целым числом")
        return
    if n < 1:
        tk.messagebox.showerror("Ошибка", "n должно быть целым числом болшим или равным 1")
        return

    try:
        words = int(e_words.get())
    except:
        tk.messagebox.showerror("Ошибка", "Число слов должно быть целым числом до от 1 до 36")
        return
    if words < 1:
        tk.messagebox.showerror("Ошибка", "Число слов должно быть целым числом болшим или равным 1")
        return

    if n * words > 4000001:
        tk.messagebox.showerror("Ошибка", "n и/или words слишком большая величина")
        return

    file_path = e_path.get()
    #is Path field empty?
    if os.path.isdir(file_path):
        if file_path[-2]==':':
            #in case of "D:\" etc.
            file_path += "DSM.xlsx"
        else:
            file_path += "\DSM.xlsx"
    else:
        tk.messagebox.showerror("Ошибка", "Указан неверный путь сохранения")
        return

    # Does file already exist?
    if os.path.isfile(file_path):
        if not tk.messagebox.askokcancel("Предупреждение","DSM.xlsx уже существует, перезаписать файл?"):
            print("closed")
            return
        try:
            #is file available?
            os.rename(file_path,file_path)
        except:
            tk.messagebox.showerror("Ошибка", "Закройте DSM.xlsx")
            return
    # Getting data and handling errors

    try:
        workbook = xl.Workbook(file_path)
    except:
        tk.messagebox.showerror("Ошибка", "Возникла ошибки при создании таблицы.")

    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True,
                                'align': 'center',
                                'border': 2})

    center = workbook.add_format({'align': 'center',
                                  'border': 2})

    def calcAvg(row_s,column_s,row_e,column_e):
        """Returning .xlsx formula for ariphmetic average"""
        # chr(65) - "A",chr(66) - "B" ...
        return "=AVERAGE(" + chr(65+column_s) + str(row_s) + ":" + chr(65+column_e) + str(row_e) + ")"

    def DrawTable(row, column,mode=1):
        # adding header
        worksheet.write_row(row, column, header, bold)
        worksheet.set_column(column, column, CELL_WIDTH)
        worksheet.set_column(column+6, column+7,CELL_WIDTH)

        # adding random data
        # setting random seed from time(NULL)
        SetSeed()
        for i in range(1, 6):
            list = GetList(Kn, n*words,mode)
            worksheet.write_column(row + 1, column + i, list, center)

        #adding symbols
        for item in alphabet:
            row += 1
            if row>Kn:
                break
            worksheet.write(row, column, item, center)

            # formulas to calculate probabilities

            p3 = calcAvg(row+1,column+1,row+1,column+3) +"/"+  str(words * n)
            p5 = calcAvg(row+1,column+1,row+1,column+5)  +"/"+ str(words * n)

            worksheet.write(row, column + 6, p3,center)
            worksheet.write(row, column + 7, p5,center)

        if Kn == 36:
            #handling edge case
            row += 1

        #adding Entropy footer
        worksheet.write(row,column,"Энтропия (апостр.)",bold)
        worksheet.write(row+1,column,"Энтропия (апр.)",bold)

        #Calculated enthropy
        worksheet.merge_range(row+1,column+1,row+1,column+7,H[mode-1],center)

        #Average enthropy
        worksheet.write(row, column + 6, calcAvg(row + 1,column + 1,row + 1,column + 3), center)
        worksheet.write(row, column + 7, calcAvg(row + 1,column + 1,row + 1,column + 5), center)

        # adding footer
        row = row + 3
        worksheet.write(row, column, "Режим", bold)

        #Source rate and speed
        worksheet.merge_range(row, column + 1, row, column + 3, "Производительность (бит/с)", bold)
        worksheet.merge_range(row, column + 4, row, column + 6, "Техническая скорость (симв/c)",bold)
        worksheet.write(row + 1, column, "Slow (250 мс)", center)
        worksheet.write(row + 2, column, "Fast (125 мс)", center)

        #H'=H/(T/(Kn*n))
        dH_slow = "=ROUNDDOWN(" + chr(65 + column + 7) + str(row - 2) + "/" + str(T_SLOW/(Kn*n)) + ",0)"
        dH_fast = "=ROUNDDOWN(" + chr(65 + column + 7) + str(row - 2) + "/" + str(T_FAST/(Kn*n)) + ",0)"
        worksheet.merge_range(row + 1, column + 1, row + 1, column + 3, dH_slow, center)
        worksheet.merge_range(row + 2, column + 1, row + 2, column + 3, dH_fast, center)
        #R=(Kn*n)/T
        R_slow = "=ROUNDDOWN(" +str((Kn*n) / T_SLOW)+ ",0)"
        R_fast = "=ROUNDDOWN(" +str((Kn*n) / T_FAST)+ ",0)"
        worksheet.merge_range(row + 1, column + 4, row + 1, column + 6, R_slow, center)
        worksheet.merge_range(row + 2, column + 4, row + 2, column + 6, R_fast, center)

        #Adding column chart
        column_chart = workbook.add_chart({'type': 'column'})

        # Adding series for chart
        column_chart.set_y_axis({'min': 0, })
        # chr(65) - "A",chr(66) - "B" ...
        column_chart.add_series({
            'name': '3 эксперемента',
            'categories': '=Sheet1!'+chr(65+column)+'2:'+chr(65+column)+str(Kn+1),
            'values': '=Sheet1!'+chr(65+column+6)+'2:'+chr(65+column+6)+str(Kn+1),
        })
        column_chart.add_series({
            'name': '5 эксперементов',
            'categories': '=Sheet1!' + chr(65 + column) + '2:' + chr(65 + column) + str(Kn + 1),
            'values': '=Sheet1!' + chr(65 + column + 7) + '2:' + chr(65 + column + 7) + str(Kn + 1),
        })

        worksheet.insert_chart(row+4,column, column_chart)

    DrawTable(0,0,mode=1)
    DrawTable(0, 9,mode=2)
    DrawTable(0, 18,mode=3)

    workbook.close()
    tk.messagebox.showinfo("Успех", "Таблица успешно сгенерирована!")
    print("Table is generated in directory", file_path)

#creating window and GUI
window = tk.Tk()

window.title("DSM")
window.geometry(WIN_SIZE)


l_Kn = tk.Label(window, text = "Объём алфавита Kn:")
l_n = tk.Label(window, text = "Длина слов n:")
l_words = tk.Label(window, text = "Число слов в предложении:")

e_Kn = tk.Entry(window)
e_n = tk.Entry(window)
e_words = tk.Entry(window)

l_H = tk.Label(window, text = "Априорная Энтропия")
l_H1 = tk.Label(window, text = "Задание №1:")
l_H2 = tk.Label(window, text = "Задание №2:")
l_H3 = tk.Label(window, text = "Задание №3:")

e_H1 = tk.Entry(window)
e_H2 = tk.Entry(window)
e_H3 = tk.Entry(window)

b_generate = tk.Button(window, text = "Generate", command = GenerateXLS)

l_path = tk.Label(window, text = "Путь сохранения:")
e_path = tk.Entry(window)
b_path = tk.Button(window, text = "View", command = lambda : browseDirectories(e_path))

#Placing elements of GUI at their places
l_Kn.grid(row=0, column=0, pady=2, sticky=tk.E)
l_n.grid(row=1, column=0, pady=2, sticky=tk.E)
l_words.grid(row=2, column=0, pady=2, sticky=tk.E)

e_Kn.grid(row=0, column=1, pady=2)
e_n.grid(row=1, column=1, pady=2)
e_words.grid(row=2, column=1, pady=2)

l_H.grid(row=3, column=0, columnspan=3, pady=2)
l_H1.grid(row=4, column=0, pady=2, sticky=tk.E)
l_H2.grid(row=5, column=0, pady=2, sticky=tk.E)
l_H3.grid(row=6, column=0, pady=2, sticky=tk.E)

e_H1.grid(row=4, column=1, pady=2)
e_H2.grid(row=5, column=1, pady=2)
e_H3.grid(row=6, column=1, pady=2)

b_generate.grid(row=7, column=1, pady=2, sticky=tk.W)

l_path.grid(row=8, column=0, pady=2, sticky=tk.E)
e_path.grid(row=8, column=1, pady=2)
b_path.grid(row=8, column=2, pady=2)

l_copyright = tk.Label(window, text = "НГТУ 2021: гр. 18-ССК Макаров Н.В. Капля В.В. Парамонов А.С.")
l_copyright.grid(row=9, column=0, columnspan=3, pady=2, sticky=tk.W)

#setting icon
try:
    window.iconbitmap(icon)
except:
    pass

window.mainloop()
#creating window and GUI
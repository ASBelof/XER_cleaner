'''
Программа для обработки xer-файлов.
    1. Импорт xer-файла.
    2. Очистка от таблиц POBS и RISCTYPE и сохранение нового xer.
    3. Сохранение xer в формате excel.
    4. Выбор таблиц для дальнейшего сохранения их в новый xer или excel
    
    последнее сохранение - 29.04.2025
    
    ИДЕИ:
        - доработать функции класса, чтобы они работали по структуре data2
        
'''


import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from pandas.io.excel import ExcelWriter

#************************
# класс - НАЧАЛО
#

class XerFile:
    selected_table_list = []
    file_path = ''
    def __init__(self, file_path, data, data2, table_list, columns):
        self.file_path = file_path
        self.data = data
        self.data2 = data2
        self.table_list = table_list
        self.columns = columns
        
    def __repr__(self):
        return f'Файл {self.file_path}\nКоличество таблиц - {len(self.table_list)}'
    
    
    def clean_xer(self):
        
        if self.file_path != "":
            
            # имя исходного xer-файла
            xer_file_name = os.path.splitext(os.path.basename(self.file_path))[0]
            print('имя исходного xer-файла - ', xer_file_name)
            
            # путь к исходному xer-файлу
            path_to_xer = os.path.dirname(self.file_path)
            print (path_to_xer)
            
            # полный путь к новому xer-файлу
            path_to_new_xer = os.path.join(path_to_xer, xer_file_name + '_NEW.xer')
            print(path_to_new_xer)
            
            with open(path_to_new_xer, 'w', encoding='cp1251', errors = 'ignore') as f:
                # запись первой строки
                f.write(self.data[0])
                
                # запись основных данных
                for table_name, rows in self.data2.items():
                    if table_name in self.selected_table_list:
                        if not rows: # удалить???
                            continue
                            
                        # Записываем заголовок таблицы (с табуляцией)
                        f.write("%T\t" + table_name + "\n")
                        
                        # Получаем все колонки
                        # all_columns = set()
                        # for row in rows:
                        #     all_columns.update(row.keys())
                        columns = self.columns.get(table_name)
                        
                        # Записываем названия столбцов (с табуляцией)
                        f.write("%F\t" + "\t".join(columns) + "\n")
                        
                        # Записываем строки данных (с табуляцией)
                        for row in rows:
                            values = [str(row.get(col, '')) for col in columns]
                            f.write("%R\t" + "\t".join(values) + "\n")
                        
                # запись последней строки
                f.write('%E')
                                
                tk.messagebox.showinfo("Результат",('Готово!!!\nВ той же папке создан xer-файл с индексом "NEW"'))
        else:
            tk.messagebox.showinfo("Ошибка!!!",('XER-файл не выбран!!!'))
            
            
    def xer_2_excel(self):
        if self.file_path != "":
            xer_file = self.file_path
            xer_file_name = os.path.splitext(os.path.basename(xer_file))[0]
            print('xer_file_name -', xer_file_name)
        
            path_to_xer = os.path.dirname(xer_file)
            print (path_to_xer)
        
            path_to_excel = os.path.join(path_to_xer, xer_file_name + '.xlsx')
            
            # path_to_excel = tk.filedialog.asksaveasfilename(filetypes = [('Excel','*.xlsx')])
            print('path_to_excel -', path_to_excel)

            # удаление файла, если он существует
            if os.path.exists(path_to_excel):
                print('Удаление существующего файла с именем', xer_file_name + '.xlsx')
                os.remove(path_to_excel)
            else: None
            
            max_rows = 1048570
            table_name = None
            fields = None
            rows = []
            list_of_series = []

            # def fill_excel_with_data (data, columns, sheet):
                                
            #     df = pd.DataFrame(data)
            #     df.columns = columns
            #     print(f"Кол-во строк в таблице {sheet} - {len(df)}")
            #     if len(df) < max_rows:
            #         df.to_excel(writer, sheet_name=sheet, index = False)
            #     else:
            #         list_qty = len(df)//max_rows
            #         for i in range(1,list_qty+2):
            #             df.iloc[max_rows*(i-1):max_rows*i].to_excel(writer, sheet_name=sheet+"_"+str(i), index = False)
                
            with ExcelWriter(path_to_excel, engine = 'openpyxl', mode="a" if os.path.exists(path_to_excel) else "w") as writer:
                print('проверка строк...')
                
                for table_name, rows in self.data2.items():
                    if table_name in self.selected_table_list:
                        
                        # Получаем порядок столбцов для текущей таблицы
                        columns = self.columns.get(table_name)
                        
                        # Создаем DataFrame с нужным порядком столбцов
                        df = pd.DataFrame(rows, columns=columns)
                        
                        # Очистка от недопустимых символов?????????????????????????????????????????????????
                        for col in df.columns:
                            if df[col].dtype == object:
                                df[col] = df[col].apply(lambda x: x.replace('\x00', '') if isinstance(x, str) else x)
                        
                        # Разбиваем на части если превышено max_rows
                        total_rows = len(df)
                        if total_rows <= max_rows:
                            df.to_excel(writer, sheet_name=table_name[:31], index=False)
                        else:
                            #
                            parts = (total_rows // max_rows) + 1
                            for i in range(parts):
                                start_idx = i * max_rows
                                end_idx = (i + 1) * max_rows
                                sheet_name = f"{table_name[:28]}_{i+1}"
                                df.iloc[start_idx:end_idx].to_excel(
                                    writer, 
                                    sheet_name=sheet_name, 
                                    index=False
                                )

            tk.messagebox.showinfo("Результат",(f'Готово!!!\nВ той же папке создан excel-файл {xer_file_name + ".xlsx"}'))
        else:
            tk.messagebox.showinfo("Ошибка!!!",('XER-файл не выбран!!!'))
#
# класс - КОНЕЦ
#***************************





#**********************************************************************
# функции программы
#**********************************************************************

# функция отвечающая за выбор файла через диалоговое окно

def select_file():
    global filepath
    filepath = ''
    
    filepath = tk.filedialog.askopenfilename(
        title = 'Поиск *.xer файла',
        filetypes = [('Primavera XER-files','*.xer')]
        )
    e1.configure(state=tk.NORMAL)
    e1.delete('1.0', tk.END)
    # создание объекта
    global xer_file
    xer_file = open_file(filepath)
    xer_file.selected_table_list = xer_file.table_list

# функция отвечающая за открытие xer
# вызывается нажатием кнопки 'Выбрать файл' (btn2)
# выводит в текстовое поле программы путь к выбранному файлу, кол-во таблиц в xer, наименования таблиц с количеством записей в них
# возвращает экземпляр класса XerFile

def open_file(filepath):
    table_list = []
    if filepath != "":
        print(filepath)
        with open(filepath, 'r', encoding='cp1251', errors = 'ignore') as f:
            lines = f.readlines()
            
        tables = {}
        columns_order = {}
        current_table = None
        current_columns = None
        
        for line in lines:
            line = line.strip()

            if line.startswith('%T'):
                # Начало новой таблицы
                current_table = line.split('\t')[1]
                tables[current_table] = []
                columns_order[current_table] = []
            elif line.startswith('%F'):
                # Заголовки столбцов таблицы
                current_columns = line.split('\t')[1:]
                columns_order[current_table] = current_columns
            elif line.startswith('%R'):
                # Строка данных таблицы
                if current_table and current_columns:
                    row_data = line.split('\t')[1:]
                    row_dict = dict(zip(current_columns, row_data))
                    tables[current_table].append(row_dict)
            elif line.startswith('%E'):
                pass
                
        table_list = sorted(tables)
                
        for i in sorted(tables):
            if str(len(tables[i]))[-1] == '1' and str(len(tables[i]))[-2:] != '11':
                e1.insert(tk.END, f'\n{i} - {len(tables[i])}строкa')
            elif str(len(tables[i]))[-1] in ['2', '3', '4'] and str(len(tables[i]))[-2:] not in ['12', '13', '14']:
                e1.insert(tk.END, f'\n{i} - {len(tables[i])}строки')
            else: 
                e1.insert(tk.END, f'\n{i} - {len(tables[i])}строк')
            
        e1.insert('1.0', f'\nКол-во таблиц в выбранном файле - {len(tables)}шт.:')
        e1.insert('1.0', f'{filepath}')
        e1.configure(state=tk.DISABLED)
    else:
        e1.insert('1.0', 'Файл не выбран...')
        e1.configure(state=tk.DISABLED)
    print (table_list)
    return XerFile(filepath, lines, tables, table_list, columns_order)


# функция отвечающая за очистку xer от таблиц POBS и RISCTYPE
# вызывается нажатием кнопки 'Очистить *.xer' (btn3)
# сохраняет очищенный xer под новым именем
def clean_xer():

    xer_file.clean_xer()

# функция отвечающая за сохранение xer в формате excel
# сохраняет excel-файл в той же папке
def xer_2_excel():
    
    xer_file.xer_2_excel()



################################################################################
# вывод окна выбора таблиц
################################################################################



def insert_check_btn(text_area, list_of_tables):
    global check_btn_vars
    global check_btn_list
    
    check_btn_vars = []
    check_btn_list = []
    print(f'tbl list is {list_of_tables}')
    list_of_tables = sorted(list_of_tables)
    for index, table in enumerate(list_of_tables):
                
        var = tk.IntVar()
        var.set(1)
        check_btn_vars.append(var)
        check_btn = tk.Checkbutton(text_area, text=f'{table}',
                                   variable=var,
                                   command = None)
        check_btn_list.append(check_btn)
        if table in ['RISKTYPE', 'POBS']:
            check_btn.configure(state=tk.DISABLED, variable = var.set(0))
            check_btn_list.remove(check_btn)
        text_area.window_create('end', window=check_btn)
        text_area.insert('end', '\n')
        
    text_area.configure(state=tk.DISABLED, cursor='')


def selection_get():
    
    global selected_indexes
    global selected_tables
    selected_indexes = []
    selected_tables = []
    
    for s in check_btn_vars:
        if s.get() == 1:
            selected_indexes.append(check_btn_vars.index(s))
            selected_tables.append(xer_file.table_list[check_btn_vars.index(s)])

    print(f'\n{selected_indexes}')
    print(selected_tables)
    xer_file.selected_table_list = selected_tables
    global select_window
    select_window.destroy()
    e1.configure(state=tk.NORMAL)
    e1.insert(tk.END, f'\n\nВыбраны таблицы {selected_tables}')
    e1.configure(state=tk.DISABLED)


def select_all():
    for i in check_btn_list:
        i.select()


def deselect_all():
    for i in check_btn_list:
        i.deselect()


#######################################################################################
# ОКНО toplevel - начало
#######
def select_tbl_window():
    # global filepath

    if xer_file.file_path == '':
        tk.messagebox.showinfo("Ошибка!!!",('XER-файл не выбран!!!'))
    else:

        global select_window
        select_window = tk.Toplevel(win)
        select_window.title('Выбор таблиц')

        select_window.resizable(False, True)

        fm_1 = ttk.Frame(select_window, borderwidth=1, relief="raised")
        fm_2 = ttk.Frame(select_window, borderwidth=1, relief="raised")
        text_area = tk.Text(fm_1, bg='SystemButtonFace', width=30)
        text_area.grid(column = 0, row = 0, sticky = tk.NSEW)

        scroll_bar = tk.Scrollbar(fm_1, command=text_area.yview)
        scroll_bar.grid(column = 1, row = 0, sticky = tk.NSEW, pady=0)

        text_area.configure(yscrollcommand=scroll_bar.set)


        select_window.columnconfigure(0, weight=1)
        select_window.columnconfigure(1, weight=1)
        select_window.rowconfigure(0, weight=1)
        select_window.rowconfigure(1, weight=0)
        fm_1.rowconfigure(0, weight=1)
        fm_2.columnconfigure(0, weight=1)
        fm_2.columnconfigure(1, weight=1)


        btn_1 = tk.Button(fm_2, text='Выбрать', command=selection_get, relief=tk.RAISED, bd = 0.5)

        def select_fun():
            if check_all_var.get() == 1:
                for i in check_btn_list: 
                    i.select()
            elif check_all_var.get() == 0:
                for i in check_btn_list:
                    i.deselect()

        check_all_var = tk.IntVar()
        check_all_var.set(1)
        check_all = tk.Checkbutton(fm_2, text='Выбрать все', variable=check_all_var, command=select_fun)


        fm_1.grid(column=0, row=0, sticky = tk.NSEW)
        check_all.grid(column=0, row=0)
        btn_1.grid(column=1, row=0)
        fm_2.grid(column=0, row=1, sticky = tk.NSEW)
        # global insert_check_btn
        insert_check_btn(text_area, xer_file.table_list)

####
# ОКНО toplevel - конец
#############################################################################################



#**********************************************************************
# главное окно программы
#**********************************************************************


win = tk.Tk()
win.title('Обработка xer-файлов')             # Заголовок главного окна
win.config(bg="#262e3e")            # цвет фона на главном окне: bg - background вписать цвет словом или RGB-код
win.geometry('400x600+300+300')     # размеры и положение главного окна
win.resizable(True, True)         # Задается возможность изменения размеров по ширине и высоте

win.grid_columnconfigure(0, weight = 1)
win.grid_rowconfigure(0, weight = 1)


btn2 = tk.Button(
    win,
    bg='#84a724',
    fg='white',
    text = 'Выбрать файл',
    command = select_file,
    relief = tk.RAISED, bd = 0.5, font='arial'
    )

btn3 = tk.Button(
    win,
    text = 'Очистить Xer-файл',
    command = clean_xer,
    bg = '#87651D',
    fg='white',
    # width = 40,
    relief = tk.RAISED, bd = 0.5, font='arial'
    )

btn4 = tk.Button(
    win,
    text = 'Экспорт в excel',
    command = xer_2_excel,
    bg = '#87651D',
    fg='white',
    # width = 40,
    relief = tk.RAISED, bd = 0.5, font='arial'
    )

# виджет вывода текстовой информации
e1 = tk.Text(
    # width=47,
    height=12,
    bg = '#ebf1f1',
    font=('Arial', 8),
    state = tk.DISABLED,
    wrap=tk.CHAR,
    relief = tk.RAISED,
    bd = 0.5
    )

scrollbar = ttk.Scrollbar(orient="vertical", command=e1.yview)
scrollbar.grid(column = 1, row = 0, sticky = tk.NS)
e1["yscrollcommand"]=scrollbar.set

btn5 = tk.Button(
    win,
    text = 'Выбрать таблицы',
    command = select_tbl_window,
    bg = '#309054',
    fg='white',
    width = 40,
    relief = tk.RAISED, bd = 0.5
    )

e1.grid(column = 0, row = 0, sticky = tk.NSEW)
btn2.grid(columnspan=2, sticky=tk.EW)
btn3.grid(columnspan=2, sticky=tk.EW)

btn4.grid(columnspan=2, sticky=tk.EW)
btn5.grid(columnspan=2, sticky=tk.EW)

win.mainloop()
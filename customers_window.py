import tkinter as tk
from tkinter.messagebox import showerror
from mwx import MWX
from table import Table


class CustWin:
    def __init__(self, master):
        self.master = master
        self.master.title("Контрагенты")
        self.book = MWX()

        self.active_row = 0
        self.frame_table = tk.Frame(self.master)
        self.table = Table(self.frame_table,
                           headings=('ИНН', 'КПП', 'Наименование'),
                           rows=self.book.get_str_list(
                               name_list='counterparty',
                               columns=[{'position': 1, 'format': '', 'trans': ''},
                                        {'position': 9, 'format': '', 'trans': ''},
                                        {'position': 2, 'format': '', 'trans': ''}]),
                           width_column=(80, 80, 450))
        self.table.table.bind('<<TreeviewSelect>>', self.on_select)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)

        self.frame_data = tk.Frame(self.master)
        self.frame_line0 = tk.Frame(self.master)
        self.frame_line1 = tk.Frame(self.master)
        self.frame_line2 = tk.Frame(self.master)
        self.frame_line3 = tk.Frame(self.master)
        self.lbl_inn = tk.Label(self.frame_line0, text="ИНН: ")
        self.lbl_inn.pack(side=tk.LEFT)
        self.ent_inn = tk.Entry(self.frame_line0, width=19)
        self.ent_inn.pack(side=tk.LEFT)
        self.lbl_kpp = tk.Label(self.frame_line0, text="КПП: ")
        self.lbl_kpp.pack(side=tk.LEFT)
        self.ent_kpp = tk.Entry(self.frame_line0, width=18)
        self.ent_kpp.pack(side=tk.LEFT)
        self.lbl_okpo = tk.Label(self.frame_line0, text="ОКПО: ")
        self.lbl_okpo.pack(side=tk.LEFT)
        self.ent_okpo = tk.Entry(self.frame_line0, width=18)
        self.ent_okpo.pack(side=tk.LEFT)
        self.lbl_tel = tk.Label(self.frame_line0, text="тел.: ")
        self.lbl_tel.pack(side=tk.LEFT)
        self.ent_tel = tk.Entry(self.frame_line0, width=22)
        self.ent_tel.pack(side=tk.LEFT)

        self.lbl_name = tk.Label(self.frame_line1, text="Наименование:  ")
        self.lbl_name.pack(side=tk.LEFT)
        self.ent_name = tk.Entry(self.frame_line1, width=88)
        self.ent_name.pack(side=tk.LEFT)

        self.lbl_adr = tk.Label(self.frame_line2, text="Адрес: ")
        self.lbl_adr.pack(side=tk.LEFT)
        self.ent_adr = tk.Entry(self.frame_line2, width=97)
        self.ent_adr.pack(side=tk.LEFT)

        self.lbl_bank = tk.Label(self.frame_line3, text="Банк: ")
        self.lbl_bank.pack(side=tk.LEFT)
        self.ent_bank = tk.Entry(self.frame_line3, width=30)
        self.ent_bank.pack(side=tk.LEFT)
        self.lbl_bik = tk.Label(self.frame_line3, text="БИК: ")
        self.lbl_bik.pack(side=tk.LEFT)
        self.ent_bik = tk.Entry(self.frame_line3, width=8)
        self.ent_bik.pack(side=tk.LEFT)
        self.lbl_rsc = tk.Label(self.frame_line3, text="р/сч: ")
        self.lbl_rsc.pack(side=tk.LEFT)
        self.ent_rsc = tk.Entry(self.frame_line3, width=20)
        self.ent_rsc.pack(side=tk.LEFT)
        self.lbl_ksc = tk.Label(self.frame_line3, text="к/сч: ")
        self.lbl_ksc.pack(side=tk.LEFT)
        self.ent_ksc = tk.Entry(self.frame_line3, width=20)
        self.ent_ksc.pack(side=tk.LEFT)

        self.frame_key = tk.Frame(self.master)
        self.button0 = tk.Button(self.frame_key, text='Добавить', width=15, command=self.add_cust)
        self.button0.pack(side=tk.RIGHT)
        self.button1 = tk.Button(self.frame_key, text='Сохранить', width=15, command=self.save_cust)
        self.button1.pack(side=tk.RIGHT)
        self.button2 = tk.Button(self.frame_key, text='Удалить', width=15, command=self.dell_cust)
        self.button2.pack(side=tk.RIGHT)
        self.quitButton = tk.Button(self.frame_key, text='Выход', width=15, command=self.close_windows)
        self.quitButton.pack(side=tk.RIGHT)
        self.frame_table.pack()
        self.frame_line0.pack(fill=tk.X)
        self.frame_line1.pack(fill=tk.X)
        self.frame_line2.pack(fill=tk.X)
        self.frame_line3.pack(fill=tk.X)
        self.frame_data.pack(fill=tk.X)
        self.frame_key.pack(fill=tk.X)

    def clear_row(self):
        self.ent_inn.delete("0", tk.END)
        self.ent_kpp.delete("0", tk.END)
        self.ent_okpo.delete("0", tk.END)
        self.ent_tel.delete("0", tk.END)
        self.ent_name.delete("0", tk.END)
        self.ent_adr.delete("0", tk.END)
        self.ent_bank.delete("0", tk.END)
        self.ent_bik.delete("0", tk.END)
        self.ent_rsc.delete("0", tk.END)
        self.ent_ksc.delete("0", tk.END)

    def up_table(self, name_sheet):
        self.active_row = 0
        self.clear_row()
        rows = self.book.get_str_list(name_list=name_sheet,
                                      columns=[{'position': 1, 'format': '', 'trans': ''},
                                               {'position': 9, 'format': '', 'trans': ''},
                                               {'position': 2, 'format': '', 'trans': ''}])
        self.table.update(rows)

    def save_data(self):
        name_sheet = 'counterparty'
        list_data = [self.ent_inn.get(),
                     self.ent_name.get(),
                     self.ent_tel.get(),
                     self.ent_adr.get(),
                     self.ent_bik.get(),
                     self.ent_bank.get(),
                     self.ent_rsc.get(),
                     self.ent_ksc.get(),
                     self.ent_kpp.get(),
                     self.ent_okpo.get()]
        if self.book.test_double(name_sheet, self.active_row, [[1, list_data[0]], [9, list_data[9]]]) == 0:
            self.book.save_list(name_sheet, self.active_row, list_data)
            self.up_table(self, name_sheet)
        else:
            showerror(title='Ошибка сохранение',
                      message=f'Контрагент с ИНН {list_data[0]} и КПП {list_data[8]} уже заведен!')

    def add_cust(self):
        self.active_row = 0
        self.save_data()

    def save_cust(self):
        if self.active_row > 1:
            self.save_data()

    def dell_cust(self, name_sheet='counterparty'):
        if self.active_row > 0:
            self.book.del_row(name_sheet, self.active_row)
            self.up_table(name_sheet)

    def close_windows(self):
        self.master.destroy()

    def on_select(self, event):
        page = 'counterparty'
        item = None
        for selected_item in self.table.table.selection():
            item = self.table.table.item(selected_item)

        if item is None:
            return

        list_search = [[1, str(item['values'][0])], [9, str(item['values'][1])]]
        self.active_row = self.book.search_row(page, list_search)

        if self.active_row > 1:
            self.clear_row()
            self.ent_inn.insert(0, self.book.get_value(page, self.active_row, 1))
            self.ent_kpp.insert(0, self.book.get_value(page, self.active_row, 9))
            self.ent_okpo.insert(0, self.book.get_value(page, self.active_row, 10))
            self.ent_tel.insert(0, self.book.get_value(page, self.active_row, 3))
            self.ent_name.insert(0, self.book.get_value(page, self.active_row, 2))
            self.ent_adr.insert(0, self.book.get_value(page, self.active_row, 4))
            self.ent_bank.insert(0, self.book.get_value(page, self.active_row, 6))
            self.ent_bik.insert(0, self.book.get_value(page, self.active_row, 5))
            self.ent_rsc.insert(0, self.book.get_value(page, self.active_row, 7))
            self.ent_ksc.insert(0, self.book.get_value(page, self.active_row, 8))

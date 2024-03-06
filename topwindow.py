import tkinter as tk
from tkinter.messagebox import showerror
from mwx import MWX
from customers_window import CustWin
from table import Table


class TopWin:
    def __init__(self, master):
        self.master = master
        self.master.title("Путевые документы ЛНП")
        self.frame = tk.Frame(self.master)
        self.book = MWX()
        self.active_row = 0
        self.table = Table(self.frame,
                           headings=('№', 'Дата', 'Контрагент'),
                           rows=self.book.get_str_list(
                               name_list='Docs',
                               columns=[{'position': 1, 'format': '', 'trans': ''},
                                        {'position': 4, 'format': '', 'trans': ''},
                                        {'position': 3, 'format': '', 'trans': 'counterparty'}]),
                           width_column=(50, 75, 300))
        self.table.table.bind('<<TreeviewSelect>>', self.on_select)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)

        self.frame_data = tk.Frame(self.master)
        self.frame_line0 = tk.Frame(self.frame_data)
        self.frame_line1 = tk.Frame(self.frame_data)
        self.frame_line2 = tk.Frame(self.frame_data)
        self.frame_line3 = tk.Frame(self.frame_data)
        self.frame_line4 = tk.Frame(self.frame_data)
        self.frame_line5 = tk.Frame(self.frame_data)

        # ДАННЫЕ О ДОКУМЕНТЕ
        self.lbl_nod = tk.Label(self.frame_line0, text="№ док.: ")
        self.ent_nod = tk.Entry(self.frame_line0, width=12)
        self.lbl_dod = tk.Label(self.frame_line0, text="Дата: ")
        self.ent_dod = tk.Entry(self.frame_line0, width=14)
        self.lbl_vrd = tk.Label(self.frame_line0, text="Время")
        self.ent_vrd = tk.Entry(self.frame_line0, width=12)
        self.lbl_sf_ = tk.Label(self.frame_line0, text="Счет-фактура")
        self.ent_sf_ = tk.Entry(self.frame_line0, width=14)

        self.lbl_nod.pack(side=tk.LEFT)
        self.ent_nod.pack(side=tk.LEFT)
        self.lbl_dod.pack(side=tk.LEFT)
        self.ent_dod.pack(side=tk.LEFT)
        self.lbl_vrd.pack(side=tk.LEFT)
        self.ent_vrd.pack(side=tk.LEFT)
        self.lbl_sf_.pack(side=tk.LEFT)
        self.ent_sf_.pack(side=tk.LEFT)

        # ДАННЫЕ О КОНТРАГЕНТЕ
        self.lbl_inn = tk.Label(self.frame_line1, text="ИНН:    ")
        self.ent_inn = tk.Entry(self.frame_line1, width=19)
        self.ent_inn.bind('<FocusOut>', self.prn_lbl)
        self.lbl_kpp = tk.Label(self.frame_line1, text="КПП: ")
        self.ent_kpp = tk.Entry(self.frame_line1, width=18)
        self.ent_kpp.bind('<FocusOut>', self.prn_lbl)
        self.lbl_nam = tk.Label(self.frame_line1, text="", width=34)

        self.lbl_inn.pack(side=tk.LEFT)
        self.ent_inn.pack(side=tk.LEFT)
        self.lbl_kpp.pack(side=tk.LEFT)
        self.ent_kpp.pack(side=tk.LEFT)
        self.lbl_nam.pack(side=tk.LEFT)

        # ДАННЫЕ О ТОВАРЕ
        self.lbl_art = tk.Label(self.frame_line2, text="Артикул: ")
        self.ent_art = tk.Entry(self.frame_line2, width=8)
        self.lbl_tov = tk.Label(self.frame_line2, text="")
        self.lbl_cos = tk.Label(self.frame_line2, text="Цена: ")
        self.ent_cos = tk.Entry(self.frame_line2, width=8)
        self.lbl_kol = tk.Label(self.frame_line2, text="Количество: ")
        self.ent_kol = tk.Entry(self.frame_line2, width=10)
        self.lbl_mes = tk.Label(self.frame_line2, text="Мест: ")
        self.ent_mes = tk.Entry(self.frame_line2, width=6)
        self.lbl_nds = tk.Label(self.frame_line2, text="НДС: ")
        self.ent_nds = tk.Entry(self.frame_line2, width=12)

        self.lbl_art.pack(side=tk.LEFT)
        self.ent_art.pack(side=tk.LEFT)
        self.lbl_tov.pack(side=tk.LEFT)
        self.lbl_cos.pack(side=tk.LEFT)
        self.ent_cos.pack(side=tk.LEFT)
        self.lbl_kol.pack(side=tk.LEFT)
        self.ent_kol.pack(side=tk.LEFT)
        self.lbl_mes.pack(side=tk.LEFT)
        self.ent_mes.pack(side=tk.LEFT)
        self.lbl_nds.pack(side=tk.LEFT)
        self.ent_nds.pack(side=tk.LEFT)

        # ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ О ТОВАРЕ
        self.lbl_pas = tk.Label(self.frame_line3, text="Паспорт: ")
        self.ent_pas = tk.Entry(self.frame_line3, width=8)
        self.lbl_plo = tk.Label(self.frame_line3, text="Плотность: ")
        self.ent_plo = tk.Entry(self.frame_line3, width=8)
        self.lbl_tem = tk.Label(self.frame_line3, text="t: ")
        self.ent_tem = tk.Entry(self.frame_line3, width=6)
        self.lbl_ddg = tk.Label(self.frame_line3, text="Дата дог.: ")
        self.ent_ddg = tk.Entry(self.frame_line3, width=10)
        self.lbl_ndg = tk.Label(self.frame_line3, text="№ дог.: ")
        self.ent_ndg = tk.Entry(self.frame_line3, width=12)

        self.lbl_pas.pack(side=tk.LEFT)
        self.ent_pas.pack(side=tk.LEFT)
        self.lbl_plo.pack(side=tk.LEFT)
        self.ent_plo.pack(side=tk.LEFT)
        self.lbl_tem.pack(side=tk.LEFT)
        self.ent_tem.pack(side=tk.LEFT)
        self.lbl_ddg.pack(side=tk.LEFT)
        self.ent_ddg.pack(side=tk.LEFT)
        self.lbl_ndg.pack(side=tk.LEFT)
        self.ent_ndg.pack(side=tk.LEFT)

        # ДАННЫЕ О ТАРНСПОРТЕ
        self.lbl_ama = tk.Label(self.frame_line4, text="Авто марка: ")
        self.ent_ama = tk.Entry(self.frame_line4, width=17)
        self.lbl_ano = tk.Label(self.frame_line4, text="гос.№: ")
        self.ent_ano = tk.Entry(self.frame_line4, width=10)
        self.lbl_vod = tk.Label(self.frame_line4, text="Водитель: ")
        self.ent_vod = tk.Entry(self.frame_line4, width=30)

        self.lbl_ama.pack(side=tk.LEFT)
        self.ent_ama.pack(side=tk.LEFT)
        self.lbl_ano.pack(side=tk.LEFT)
        self.ent_ano.pack(side=tk.LEFT)
        self.lbl_vod.pack(side=tk.LEFT)
        self.ent_vod.pack(side=tk.LEFT)

        # ДОП ИНФО
        self.lbl_pdo = tk.Label(self.frame_line5, text="Платеж.документы: ")
        self.ent_pdo = tk.Entry(self.frame_line5, width=34)
        self.lbl_plm = tk.Label(self.frame_line5, text="Пломбы: ")
        self.ent_plm = tk.Entry(self.frame_line5, width=25)

        self.lbl_pdo.pack(side=tk.LEFT)
        self.ent_pdo.pack(side=tk.LEFT)
        self.lbl_plm.pack(side=tk.LEFT)
        self.ent_plm.pack(side=tk.LEFT)

        self.frame_key = tk.Frame(self.master)
        self.button2 = tk.Button(self.frame_key, text='Контрагенты', width=12, command=self.new_cust)
        self.button2.pack(side=tk.RIGHT)
        self.button1 = tk.Button(self.frame_key, text='Новый', width=12, command=self.add_doc)
        self.button1.pack(side=tk.RIGHT)
        self.button1 = tk.Button(self.frame_key, text='Сохранить', width=12, command=self.save_doc)
        self.button1.pack(side=tk.RIGHT)
        self.button1 = tk.Button(self.frame_key, text='Удалить', width=12, command=self.dell_doc)
        self.button1.pack(side=tk.RIGHT)
        self.button0 = tk.Button(self.frame_key, text='Вывод XLS', width=12, command=self.prn_docs)
        self.button0.pack(side=tk.RIGHT)
        self.quitButton = tk.Button(self.frame_key, text='Выход', width=12, command=self.close_windows)
        self.quitButton.pack(side=tk.RIGHT)

        self.frame.pack(fill=tk.X)
        self.frame_line0.pack(fill=tk.X)
        self.frame_line1.pack(fill=tk.X)
        self.frame_line2.pack(fill=tk.X)
        self.frame_line3.pack(fill=tk.X)
        self.frame_line4.pack(fill=tk.X)
        self.frame_line5.pack(fill=tk.X)

        self.frame_data.pack(fill=tk.X)
        self.frame_key.pack(fill=tk.X)

    def prn_lbl_real(self):
        list_search = [[1, self.ent_inn.get()], [9, self.ent_kpp.get()]]
        name = self.book.get_name('counterparty', list_search, "Анонимус")
        self.lbl_nam.configure(text=name)

    def prn_lbl(self, event):
        self.prn_lbl_real()

    def clear_row(self):
        self.ent_nod.delete("0", tk.END)
        self.ent_dod.delete("0", tk.END)
        self.ent_vrd.delete("0", tk.END)
        self.ent_sf_.delete("0", tk.END)
        self.ent_inn.delete("0", tk.END)
        self.ent_kpp.delete("0", tk.END)
        self.ent_art.delete("0", tk.END)
        self.ent_cos.delete("0", tk.END)
        self.ent_kol.delete("0", tk.END)
        self.ent_mes.delete("0", tk.END)
        self.ent_nds.delete("0", tk.END)
        self.ent_pas.delete("0", tk.END)
        self.ent_plo.delete("0", tk.END)
        self.ent_tem.delete("0", tk.END)
        self.ent_ddg.delete("0", tk.END)
        self.ent_ndg.delete("0", tk.END)
        self.ent_ama.delete("0", tk.END)
        self.ent_ano.delete("0", tk.END)
        self.ent_vod.delete("0", tk.END)
        self.ent_pdo.delete("0", tk.END)
        self.ent_plm.delete("0", tk.END)

    def up_table(self, name_sheet):
        self.active_row = 0
        self.clear_row()
        rows = self.book.get_str_list(name_list='Docs',
                                      columns=[{'position': 1, 'format': '', 'trans': ''},
                                               {'position': 4, 'format': '', 'trans': ''},
                                               {'position': 3, 'format': '', 'trans': 'counterparty'}])
        self.table.update(rows)

    def save_data(self):
        name_sheet = 'Docs'
        list_data = [self.ent_nod.get(),
                     self.ent_sf_.get(),
                     self.ent_inn.get(),
                     self.ent_dod.get(),
                     self.ent_vrd.get(),
                     self.ent_pdo.get(),
                     self.ent_pas.get(),
                     self.ent_plm.get(),
                     self.ent_vod.get(),
                     self.ent_ano.get(),
                     self.ent_ama.get(),
                     self.ent_mes.get(),
                     self.ent_nds.get(),
                     self.ent_art.get(),
                     self.ent_kol.get(),
                     self.ent_cos.get(),
                     self.ent_ndg.get(),
                     self.ent_ddg.get(),
                     self.ent_plo.get(),
                     self.ent_tem.get(), ' ',
                     self.ent_kpp.get()]
        if self.book.test_double(name_sheet, self.active_row, [[1, list_data[0]], [4, list_data[3]]]) == 0:
            self.book.save_list(name_sheet, self.active_row, list_data)
            self.up_table(name_sheet)
        else:
            showerror(title='Ошибка сохранение',
                      message=f'Документ № {list_data[0]} уже зарегистрирован!')

    def add_doc(self):
        self.active_row = 0
        self.save_data()

    def save_doc(self):
        if self.active_row > 0:
            self.save_data()

    def dell_doc(self, name_sheet='Docs'):
        if self.active_row > 1:
            self.book.del_row(name_sheet, self.active_row)
            self.up_table(name_sheet)

    def prn_docs(self):
        if self.active_row > 1:
            self.book.prn_p4(self.active_row)
            self.book.prn_t12(self.active_row)
            self.book.prn_sf(self.active_row)
        else:
            showerror(title='Ошибка вывода файлов', message=f'Не выбран документ!')

    def new_cust(self):
        CustWin(tk.Toplevel(self.master))

    def close_windows(self):
        self.master.destroy()

    def on_select(self, event):
        page = 'Docs'
        item = None

        for selected_item in self.table.table.selection():
            item = self.table.table.item(selected_item)

        if item is None:
            return

        list_search = [[1, str(item['values'][0])]]

        self.active_row = self.book.search_row(page, list_search)
        if self.active_row > 1:
            self.clear_row()
            self.ent_inn.insert(0, self.book.get_value(page, self.active_row, 3))
            self.ent_nod.insert(0, self.book.get_value(page, self.active_row, 1))
            self.ent_dod.insert(0, self.book.get_value(page, self.active_row, 4))
            self.ent_vrd.insert(0, self.book.get_value(page, self.active_row, 5))
            self.ent_sf_.insert(0, self.book.get_value(page, self.active_row, 2))
            self.ent_kpp.insert(0, self.book.get_value(page, self.active_row, 22))
            self.ent_art.insert(0, self.book.get_value(page, self.active_row, 14))
            self.ent_cos.insert(0, self.book.get_value(page, self.active_row, 16))
            self.ent_kol.insert(0, self.book.get_value(page, self.active_row, 15))
            self.ent_mes.insert(0, self.book.get_value(page, self.active_row, 12))
            self.ent_nds.insert(0, self.book.get_value(page, self.active_row, 13))
            self.ent_pas.insert(0, self.book.get_value(page, self.active_row, 7))
            self.ent_plo.insert(0, self.book.get_value(page, self.active_row, 19))
            self.ent_tem.insert(0, self.book.get_value(page, self.active_row, 20))
            self.ent_ddg.insert(0, self.book.get_value(page, self.active_row, 18))
            self.ent_ndg.insert(0, self.book.get_value(page, self.active_row, 17))
            self.ent_ama.insert(0, self.book.get_value(page, self.active_row, 11))
            self.ent_ano.insert(0, self.book.get_value(page, self.active_row, 10))
            self.ent_vod.insert(0, self.book.get_value(page, self.active_row, 9))
            self.ent_pdo.insert(0, self.book.get_value(page, self.active_row, 6))
            self.ent_plm.insert(0, self.book.get_value(page, self.active_row, 8))
            self.prn_lbl_real()

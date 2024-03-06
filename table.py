import tkinter as tk
import tkinter.ttk as ttk


class Table(tk.Frame):
    def __init__(self, parent=None, headings=tuple(), rows=None, width_column: tuple = None):
        super().__init__(parent)

        if width_column is None:
            width_column = tuple()

        self.active = 0

        self.table = ttk.Treeview(self, show="headings", selectmode="browse")
        self.table["columns"] = headings
        self.table["displaycolumns"] = headings

        for countHead in range(len(headings)):
            head = headings[countHead]
            self.table.heading(head, text=head, anchor=tk.CENTER)
            self.table.column(head, stretch=tk.NO, width=width_column[countHead])

        self.update(rows)

        self.scroll_table = ttk.Scrollbar(self, command=self.table.yview)

        self.scroll_table.pack(side=tk.RIGHT, fill=tk.Y)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)

    def update(self, rows=None) -> None:
        if rows is None:
            return
        self.table.delete(*self.table.get_children())
        for row in rows:
            self.table.insert('', tk.END, values=tuple(row))

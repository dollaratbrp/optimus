"""
This file manages all activities linked to Optimus standalone mode

Author : Nicolas Raymond
"""
from tkinter import *
import pandas as pd
from openpyxl import load_workbook

workbook_path = 'U:\LoadAutomation\Optimus\FastLoadsSKUs.xlsx'


class FastLoadsBox:

    def __init__(self, master):

        """
        Generates a GUI that allows user to valid fast load details and to start the execution
        :param master: root of the tkinter window
        """
        self.master = master
        self.master.title('Optimus FastLoads')
        self.frame = Frame(self.master, borderwidth=2, relief=RIDGE)
        self.frame.grid(row=0, column=0)

        # Sku and qty labels initialization
        self.sku_label = Label(self.frame, text='SKU', padx=50, pady=10, borderwidth=2, relief=RIDGE)
        self.qty_label = Label(self.frame, text='QTY', padx=50, pady=10, borderwidth=2, relief=RIDGE)

        # Sku and qty labels positioning
        self.sku_label.grid(row=0, column=0)
        self.qty_label.grid(row=0, column=1)

        self.labels, self.entries = self.create_labels_and_entries(self.frame, self.read_skus_and_quantities())
        self.run_button = Button(self.master, text='Run Optimus', padx=50, pady=10)
        self.run_button.grid(row=1, columnspan=2, sticky=E+W)

    @staticmethod
    def create_labels_and_entries(frame, skus_n_qty):

        # Labels and entries container
        labels = []
        entries = []

        for i in skus_n_qty.index:

            # Sku label initialization and positioning
            labels.append(Label(frame, text=str(skus_n_qty['SKU'][i]), padx=10, pady=10))
            labels[i].grid(row=i+1, column=0)

            # Qty entries initialization and positioning
            entries.append(Entry(frame, justify='center'))
            entries[i].insert(0, str(skus_n_qty['QTY'][i]))
            entries[i].grid(row=i+1, column=1)

        return labels, entries

    @staticmethod
    def read_skus_and_quantities():

        wb = load_workbook(workbook_path)
        ws = wb.active

        # We access to the table where the data on models is contained (table 0)
        table_range = ws._tables[0].ref

        # We save the data in a data frame
        skus_n_qty = build_dataframe(ws[table_range])

        return skus_n_qty


def open_fastloads_box():

    root = Tk()
    fastloadsbox = FastLoadsBox(root)
    root.mainloop()


def build_dataframe(ws):

    """
    Build a data frame (used to store models' data)

    :param ws: Worksheet or part of the worksheet use to build de pandas dataframe
    :return: Pandas dataframe

    """

    data_rows = []
    for row in ws:
        data_cols = []
        for cell in row:
            data_cols.append(cell.value)
        data_rows.append(data_cols)

    return pd.DataFrame(data=data_rows[1:], columns=data_rows[0])

if __name__ == '__main__':
    open_fastloads_box()
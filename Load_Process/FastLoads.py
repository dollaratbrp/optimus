"""
This file manages all activities linked to Optimus standalone mode

Author : Nicolas Raymond
"""
from tkinter import *
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from Import_Functions import SQLConnection

workbook_path = 'U:\LoadAutomation\Optimus\FastLoadsSKUs.xlsx'


class FastLoadsBox:

    def __init__(self, master):

        """
        Generates a GUI that allows user to valid fast load details and to start the execution
        :param master: root of the tkinter window
        """
        self.master = master
        self.master.title('Optimus FastLoads')

        # Initialization and positioning of frame that will contain sku labels
        self.frame = Frame(self.master, borderwidth=2, relief=RIDGE)
        self.frame.grid(row=0, column=0)

        # "SKU" and "QTY" labels initialization and positioning
        self.sku_label = Label(self.frame, text='SKU', padx=50, pady=10, borderwidth=2, relief=RIDGE)
        self.qty_label = Label(self.frame, text='QTY', padx=50, pady=10, borderwidth=2, relief=RIDGE)
        self.sku_label.grid(row=0, column=0)
        self.qty_label.grid(row=0, column=1)

        # All single sku labels and qty entries initialization and positioning
        self.labels, self.entries = self.create_labels_and_entries(self.frame, self.read_skus_and_quantities())

        # "Run optimus" button configurations
        self.run_button = Button(self.master, text='Run Optimus', padx=50, pady=10, command=self.run_optimus)
        self.run_button.grid(row=1, columnspan=2, sticky=E+W)

    @staticmethod
    def create_labels_and_entries(frame, skus_n_qty):
        """
        Creates and places wisely sku labels and qty entries in our frame
        :param frame: frame widget of our FastLoadsBox
        :param skus_n_qty: pandas dataframe containing SKUs and quantities associated to each
        :return: list of labels and list of entries
        """

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
        """
        Reads SKUs and quantities in the Excel table for fast loads
        :return: pandas dataframe containing SKUs and quantities associated to each
        """

        wb = load_workbook(workbook_path)
        ws = wb.active

        # We access to the table where the data on models is contained (table 0)
        table_range = ws._tables[0].ref

        # We save the data in a data frame
        skus_n_qty = build_dataframe(ws[table_range])

        return skus_n_qty

    def run_optimus(self):
        """
        Run Optimus program on our selection of SKUs
        :return:
        """
        skus_list, qty_list = self.save_skus_and_quantities()
        complete_dataframe = self.get_complete_dataframe(skus_list, qty_list)
        self.split_dataframes(complete_dataframe)

        pass

    def save_skus_and_quantities(self):
        """
        Save definitive lists of SKUs and quantities to use once user push "Run Optimus".
        :return: list of skus and list of qty
        """

        skus_list = []  # To find data in sql
        qty_list = []   # To use as column in pandas dataframe

        # We save each sku in a list if their quantity is greater than 0
        for i in range(len(self.labels)):
            qty = int(self.entries[i].get())
            if qty > 0:
                skus_list.append(self.labels[i]['text'])
                qty_list.append(qty)

        return skus_list, qty_list

    @staticmethod
    def get_complete_dataframe(skus_list, qty_list):
        """
        Retrieves size_code and dimensions associated with our SKUs
        plus : 'NBR_PER_CRATE', 'STACK_LIMIT' and 'OVERHANG'

        :param skus_list: list of SKUs used in sql query to get the data
        :param qty_list: list of quantities associated to each SKU
        :return: pandas dataframe
        """
        # Initialization of column names for data that will be sent to LoadBuilder
        columns = ['QTY', 'SKU', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']

        # Connection to SQL database that contains data needed
        sql_connect = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_0_MD_D_MATERIAL')

        # Writing of our query
        sql_query = """ SELECT RTRIM(a.Material_Number)
        ,RTRIM(a.Size_Dimensions)
        ,CONVERT(int, CEILING(b.Length))
        ,CONVERT(int, CEILING(b.Width))
        ,CONVERT(int, CEILING(b.HEIGHT))
        ,CASE WHEN c.SUB_DIVISION = 'RYKER' THEN 2 ELSE 1 END
        ,CASE WHEN b.HEIGHT = 0 THEN 1 ELSE CONVERT(int, FLOOR(105/b.HEIGHT)) END
        ,CASE WHEN (CASE WHEN b.HEIGHT = 0 THEN 1 ELSE FLOOR(105/b.HEIGHT) END) = 1 THEN 0 ELSE 1 END
        FROM OTD_0_MD_D_MATERIAL as c LEFT JOIN MasterData.dbo.MD_MARA as a
        on c.MATERIAL_NUMBER = a.Material_Number LEFT JOIN MasterData.dbo.MD_MARA as b
        on b.Material_Number = a.Ref_Mat_Packed_In_Same_Way
        WHERE a.Material_Number in """ + str(tuple(skus_list))

        # Retrieve the data
        data = sql_connect.GetSQLData(sql_query)

        # Add each qty to the good line of data
        for line in data:
            index_of_qty = skus_list.index(line[0])
            line.insert(0, qty_list[index_of_qty])

        return pd.DataFrame(data=data, columns=columns)

    @staticmethod
    def split_dataframes(complete_dataframe):
        """
        Takes the complete dataframe and split it in two
        First : [QTY | SKU |  MODEL (SIZE_CODE)] to keep track of link between SKUs and size_code
        Second : [QTY | MODEL (SIZE_CODE) | LENGTH | WIDTH | HEIGHT | NUMBER_PER_CRATE | STACK_LIMIT | OVERHANG ]
        The second will be group by MODEL

        :return: two pandas data frames
        """

        # We extract both dataframes needed by making copy of some parts on complete dataframe
        first_df = complete_dataframe[['QTY', 'SKU', 'MODEL']].copy()
        second_df = complete_dataframe[['QTY', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE',
                                       'STACK_LIMIT', 'OVERHANG']].copy()

        # Do a groupby on second dataframe
        second_df = second_df.groupby(['MODEL', 'LENGTH', 'WIDTH', 'HEIGHT',
                                       'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']).sum().reset_index()

        return first_df, second_df


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
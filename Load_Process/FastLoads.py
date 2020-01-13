"""
This file manages all activities linked to Optimus standalone mode

Author : Nicolas Raymond
"""
from tkinter import *
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from Import_Functions import SQLConnection, savexlsxFile
from P2P_Functions import get_trailers_data
from LoadBuilder import LoadBuilder
from random import randint

workbook_path = 'U:\LoadAutomation\Optimus\FastLoadsSKUs.xlsx'

saving_path = 'U:\LoadAutomation\Optimus\\'


class FastLoadsBox:

    def __init__(self, master):

        """
        Generates a GUI that allows user to valid fast load details and to start the execution
        :param master: root of the tkinter window
        """

        # Saving of master and setting of GUI title
        self.master = master
        self.master.title('Optimus FastLoads')

        # Initialization and positioning of a frame that will contain sku labels and entries
        self.sku_frame = LabelFrame(self.master, borderwidth=2, relief=RIDGE, text='Crates')
        self.sku_frame.grid(row=0, column=0)

        # All single sku labels and qty entries initialization and positioning
        self.sku_labels, self.sku_entries = self.create_sku_labels_and_entries(self.sku_frame,
                                                                               self.read_skus_and_quantities())
        # Recuperation of trailers data
        self.trailers_data = get_trailers_data()

        # Initialization and positioning of a frame that will contain trailers labels and entries
        self.trailer_frame = LabelFrame(self.master, borderwidth=2, relief=RIDGE, text='Trailers')
        self.trailer_frame.grid(row=0, column=1)

        # All single trailer labels and qty entries initialization and positioning
        self.trailer_labels, self.trailer_qty_entries = self.create_trailer_labels_and_entries(self.trailer_frame)

        # Initialization and positioning of a frame for max loads entry
        self.max_frame = Frame(self.trailer_frame, borderwidth=2, relief=RIDGE)
        self.max_frame.grid(row=len(self.trailer_labels)+1, columnspan=2)
        self.max_label = Label(self.max_frame, text="MAX LOADS", padx=4, pady=10)
        self.max_entry = Entry(self.max_frame, justify='center', width=10)
        self.max_entry.insert(0, "∞")
        self.max_label.grid(row=0, column=0)
        self.max_entry.grid(row=0, column=1)

        # "Run optimus" button configurations
        self.run_button = Button(self.master, text='Run Optimus', padx=50, pady=10, command=self.run_optimus, bd=3,
                                 font=('TkDefaultFont', 12, 'italic'), bg='gray80')
        self.run_button.grid(row=2, columnspan=2, sticky=E+W)

        # Initialization of an empty load builder and an empty tracker attribute
        self.LoadBuilder = None
        self.tracker = None

    @staticmethod
    def create_sku_labels_and_entries(frame, skus_n_qty):
        """
        Creates and places wisely sku labels and qty entries in the frame entered as parameter
        :param frame: frame widget of our FastLoadsBox
        :param skus_n_qty: pandas dataframe containing SKUs and quantities associated to each
        :return: list of labels and list of entries
        """

        # Labels and entries container initialization
        labels = []
        entries = []

        for i in skus_n_qty.index:

            # Sku label initialization and positioning
            labels.append(Label(frame, text=str(skus_n_qty['SKU'][i]), padx=10, pady=10))
            labels[i].grid(row=i+1, column=0)

            # Qty entries initialization and positioning
            entries.append(Entry(frame, justify='center', width=10))
            entries[i].insert(0, str(skus_n_qty['QTY'][i]))
            entries[i].grid(row=i+1, column=1)

        return labels, entries

    def create_trailer_labels_and_entries(self, frame):
        """
        Creates and places wisely trailer labels and qty entries in the frame entered as parameter
        :param frame: frame widget of our FastLoadsBox
        :return: list of labels and list of entries
        """

        # Labels and entries container initialization
        labels = []
        entries = []

        for i in self.trailers_data.index:

            # Trailer label initialization and positioning
            labels.append(Label(frame, text=str(self.trailers_data['CATEGORY'][i]), padx=10, pady=10))
            labels[i].grid(row=i + 1, column=0)

            # Qty entries initialization and positioning
            entries.append(Entry(frame, justify='center', width=10))
            entries[i].insert(0, 0)
            entries[i].grid(row=i + 1, column=1)

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
        # We save definitive quantities of SKUs and trailers
        skus_list, qty_list = self.save_skus_and_quantities()
        number_of_trailers = self.save_trailers_quantities()

        if len(skus_list) != 0 and number_of_trailers != 0:

            # Building of dataframes required for the loading
            # (df1) Use to build an object that keeps track of link between SKUs and size code
            # (df2) Dataframe needed by the load builder
            complete_dataframe = self.get_complete_dataframe(skus_list, qty_list)
            df1, df2 = self.split_dataframes(complete_dataframe)

            # Initialization of the tracker (size_code dictionaries with SKUsContainer as value)
            self.tracker = self.tracker_initialization(df1)

            # Initialization of our LoadBuilder
            self.LoadBuilder = LoadBuilder(trailers_data=self.trailers_data)

            # Recuperation of the maximum of loads
            max_loads = self.max_entry.get()
            if max_loads == '∞':
                max_loads = 5000
            else:
                max_loads = int(max_loads)

            # Closing of the GUI and construction of loads
            self.master.destroy()
            size_code_used = self.LoadBuilder.build(models_data=df2, max_load=max_loads, plot_load_done=False)

            # We save all the results in a workbook at the path mentioned
            self.write_results(df2)

        pass

    def save_skus_and_quantities(self):
        """
        Save definitive lists of SKUs and quantities to use once user push "Run Optimus".
        :return: list of skus and list of qty
        """

        skus_list = []  # To find data in sql
        qty_list = []   # To use as column in pandas dataframe

        # We save each sku in a list if their quantity is greater than 0
        for i in range(len(self.sku_labels)):
            qty = int(self.sku_entries[i].get())
            if qty > 0:
                skus_list.append(self.sku_labels[i]['text'])
                qty_list.append(qty)

        return skus_list, qty_list

    def save_trailers_quantities(self):
        """
        Save definitive quantities of trailers in our trailer dataframe

        :return : total number of trailers
        """

        # Initialization of a counter of trailers
        number_of_trailers = 0

        for i in self.trailers_data.index:
            qty = int(self.trailer_qty_entries[i].get())
            self.trailers_data.at[i, 'QTY'] = qty
            number_of_trailers += qty

        return number_of_trailers

    def write_results(self, grouped_dataframe):
        """
        Saves all the results from the loading in an Excel workbook

        :param grouped_dataframe: pandas dataframe that was passed to our loadbuilder
        """

        # Initialization of a workbook
        wb = Workbook()
        fill = PatternFill(fill_type="solid", start_color="a6a6a6", end_color="a6a6a6")

        # Approved loads worksheet construction and customization
        wsApproved = wb.active
        wsApproved.title = "APPROVED"
        wsApproved.append(['LOAD_NUMBER', 'MATERIAL_NUMBER', 'QUANTITY', 'SIZE_DIMENSIONS'])

        for y in range(1, 5):
            wsApproved.cell(row=1, column=y).fill = fill

        for letter in ['A', 'B', 'C', 'D']:
            wsApproved.column_dimensions[letter].width = 20

        # Unused crates worksheet construction and customization
        wsUnused = wb.create_sheet("UNUSED")
        wsUnused.append(['MATERIAL_NUMBER', 'QUANTITY'])

        for y in range(1, 3):
            wsUnused.cell(row=1, column=y).fill = fill

        wsUnused.column_dimensions['A'].width = 15
        wsUnused.column_dimensions['B'].width = 10

        # Writing of approved loads
        self.write_approved_loads(wsApproved, grouped_dataframe)

        # Save the xlsx file
        savexlsxFile(wb=wb, path=saving_path, filename='AdHoc', Time=True)

    def write_approved_loads(self, ws, grouped_dataframe):
        """
        Writes the results for the approved loads

        :param ws: worksheet on which we write the results
        :param grouped_dataframe: pandas dataframe that was passed to our loadbuilder
        :return:
        """

        # Retrieving summary of loads done
        loads = self.LoadBuilder.get_loading_summary()
        load_number = 0

        if len(self.LoadBuilder) > 0:  # If some loads were created

            # For every line of data in our "loads" dataframe
            for i in loads.index:

                # For every load of this kind
                for iteration in range(int(loads["QTY"][i])):

                    load_number += 1

                    # For all the different size_code column in the dataframe
                    for size_code in loads.columns[4::]:

                        # if there is some quantity for this size_code
                        if loads[size_code][i] != '':

                            # We save the number of item per crate for this size_code
                            index = list(grouped_dataframe['MODEL']).index(size_code)
                            nb_per_crate = grouped_dataframe['NBR_PER_CRATE'][index]

                            # For every unit of this crate on the load
                            for unit in range(int(loads[size_code][i])):

                                # We pick a random SKU linked with the size_code
                                sku = self.tracker[size_code].random_pick()
                                ws.append([load_number, sku, nb_per_crate, size_code])

    @staticmethod
    def positive(a):
        """
        Look if the input is positive, we don't check if it's an integer because it's automatically converted as an int
        in save_skus_and_quantitites and save_trailer_quantities.
        :param a: variable
        :return: raise an Exception if the input is not positive nor integer
        """
        raise NotImplementedError

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

        # We look length of SKUs list and adapt the code
        if len(skus_list) == 1:
            end_of_query = """WHERE a.Material_Number = """ + "'" + str(skus_list[0]) + "'"

        else:
            end_of_query = """WHERE a.Material_Number in """ + str(tuple(skus_list))

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
        on b.Material_Number = a.Ref_Mat_Packed_In_Same_Way """ + end_of_query

        # Retrieve the data
        data = sql_connect.GetSQLData(sql_query)

        # Add each qty at the beginning of the good line of data
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

    @staticmethod
    def tracker_initialization(sku_dataframe):
        """
        Initializes a dictionary of size_code (key) with SKUsContainer as value
        :param sku_dataframe: pandas dataframe containing [QTY | SKU |  MODEL (SIZE_CODE)]
        :return: dictionary ("tracker")
        """
        # We count the number of unique model
        models = set(sku_dataframe['MODEL'])

        # We intialize our tracker dict and put empty SKUsContainer in it
        tracker = {}
        for model in models:
            tracker[model] = SKUsContainer()

        # We fill every SKUsContainer
        for i in sku_dataframe.index:
            tracker[sku_dataframe['MODEL'][i]].add_sku(sku_dataframe['SKU'][i], sku_dataframe['QTY'][i])

        return tracker


class SKUsContainer:

    def __init__(self):
        self.skus_dict = {}

    def add_sku(self, sku, qty):
        self.skus_dict[sku] = qty

    def __remove_sku(self, sku):
        self.skus_dict.pop(sku)

    def random_pick(self):

        # We pick a SKU randomly in our dictionary
        rand_index = randint(0, len(self.skus_dict)-1)
        skus = list(self.skus_dict.keys())
        rand_sku = skus[rand_index]

        # We decrease its quantity by one
        self.skus_dict[rand_sku] -= 1

        # We remove the sku from our dictionary if the quantity is now null
        if self.skus_dict[rand_sku] == 0:
            self.__remove_sku(rand_sku)

        # We return the SKU
        return rand_sku

    def __repr__(self):
        return str(self.skus_dict)


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
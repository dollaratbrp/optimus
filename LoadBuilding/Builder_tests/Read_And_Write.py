
"""

Created by Nicolas Raymond on 2019-05-31.

This file contains all functions used for the reading and the writing of data
associated with the load automation. Note that all interactions with Excel are managed here.

Last update : 2019-11-23
By : Nicolas Raymond

"""

import os
import sys
import subprocess
import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
from datetime import date


def read_models_data(ws):

    """

    This function read "Models" table from worksheet

    :param ws: Active Excel worksheet that contains data
    :returns: Pandas data frame with models data

    """

    # We access to the table where the data on models is contained (table 0)
    models_range = ws._tables[0].ref

    # We save the data in a data frame
    models_data = build_dataframe(ws[models_range])

    return models_data


def read_trailers_data(ws):

    """
    This function read trailers' data from worksheet

    :param ws: Active Excel worksheet that contains data
    :return: list of trailers

    """

    # We access to the table where the data on trailers is contained (table 1)
    trailers_range = ws._tables[1].ref

    # We save the data in a data frame
    trailers_data = build_dataframe(ws[trailers_range])

    return trailers_data


def show_summary(unused_list, nbr_trailers, execution_time):

    """
    Show a summary of the build status

    :param unused_list: List of models unused
    :param nbr_trailers: Number of trailers used (int)
    :param execution_time: Time of execution (in seconds)

    """
    # We open the file text (or create it if it doesn't exist)
    f = open("summary.txt", "w+")

    # We write a summary of the loading process on it
    f.write("LOADING DONE! YOU'RE AWESOME!\n\n")
    f.write("NBR OF TRAILERS : %g \n\n" % nbr_trailers)
    f.write("EXECUTION TIME : %g \n\n" % execution_time)
    f.write("UNUSED ITEMS : \n\n")
    counter = Counter(unused_list)
    for key, value in counter.items():
        line_text = str(key) + ": " + str(value)
        f.write(line_text+"\n")
    
    # We close the file
    f.close()

    # We open the .txt file
    open_file('summary.txt')


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


def open_file(filename):

    """
    Opens file according to the computer system

    :param filename: name of the file
    """
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])


def display_df(dataframe):

    """
    Displays data frame

    :param dataframe: Pandas data frame to display
    """

    fig, ax = plt.subplots()

    # hide axes
    fig.patch.set_visible(False)
    ax.axis('off')
    ax.axis('tight')

    # Create the table and set fontsize
    the_table = ax.table(cellText=dataframe.values, colLabels=dataframe.columns, loc='center')
    the_table.auto_set_font_size(False)
    the_table.set_fontsize(6)

    plt.show()


def write_summarized_data(data_frame, plant_from, plant_to, directory):

    """
    Writes a loading summary in a .xlsx file and create a folder named "P2P - <date>"
    at the directory mentioned and save the file in it.

    :param data_frame: Pandas data frame with the loading summary
    :param plant_from : String that mentions the plant from
    :param plant_to : String that mentions the plant to
    :param directory: String that mentions path where the created folder is going to be saved

    """

    # We generate today's date
    folder_date = date.today()
    folder_date = str(folder_date.strftime("%m-%d-%y"))

    # We save the folder name
    folder = folder_date + '/'

    # We save the path
    path = directory + folder

    # We create the folder in which the file will be stored (if the folder doesn't exist)
    create_folder(path)

    # We save the complete title of our future file
    title = path + "P2P from " + plant_from + " to " + plant_to + ".xlsx"

    # We initialize a "writer"
    writer = pd.ExcelWriter(title, engine='xlsxwriter')

    # We export results in the file created
    data_frame.to_excel(writer, sheet_name='Loads', index=True)

    # We initialize a workbook
    workbook = writer.book

    # We initialize a worksheet
    worksheet = writer.sheets['Loads']

    # We set the widths of the columns
    worksheet.set_column('A:B', 15)
    worksheet.set_column('B:C', 4.5)
    worksheet.set_column('C:D', 15)
    worksheet.set_column('E:AK', 4.5)

    writer.save()


def create_folder(directory):

    """
    Creates a folder with the directory mentioned

    :param directory: String that mention path where the folder is going to be created
    """
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)

    except OSError:
        print('Error while creating directory : ' + directory)

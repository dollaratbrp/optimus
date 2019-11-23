
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

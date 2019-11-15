
"""

Created by Nicolas Raymond on 2019-05-31.

This file contains all functions used for the reading and the writing of data
associated with the load automation. Note that all interactions with Excel are managed here.

Last update : 2019-10-31
By : Nicolas Raymond

"""

import LoadingObjects
import numbers
import os
import pandas as pd
import numpy as np
from collections import Counter
from datetime import date


def read_models_data(ws):

    """

    This function read "Models" table from worksheet

    :param ws: Active Excel worksheet that contains data
    :returns: A warehouse, a crate manager object and a list of model names

    """

    # Creation of a warehouse and a crate manager that will be returned by the function
    warehouse = LoadingObjects.Warehouse()
    remaining_crates = LoadingObjects.CratesManager()

    # Creation of a list that will contain model names that will also be returned
    model_names = []

    # We access to the table where the data on models is contained (table 0)
    models_range = ws._tables[0].ref

    # We save the data in a data frame
    models_data = build_dataframe(ws[models_range])

    # For all lines of the data frame
    for i in models_data.index:

        # We save the quantity of the model
        qty = models_data['QTY'][i]

        if qty > 0:

            # We save the name of the model
            model_names.append(models_data['MODEL'][i])

            # We save the stack limit
            stack_limit = models_data['STACK_LIMIT'][i]

            # We save the number of models per crate
            nbr_per_crate = models_data['NBR_PER_CRATE'][i]

            # We save the overhang permission indicator
            overhang = bool(models_data['OVERHANG'][i])

            # We compute the number of models per stack
            items_per_stack = stack_limit * nbr_per_crate

            # We compute the number of stacks that we can build
            nbr_stacks = int(np.floor(qty / items_per_stack))

            for j in range(nbr_stacks):

                # We build the stack and send it into the warehouse
                warehouse.add_stack(LoadingObjects.Stack(max(models_data['LENGTH'][i], models_data['WIDTH'][i]),
                                                         min(models_data['WIDTH'][i], models_data['LENGTH'][i]),
                                                         models_data['HEIGHT'][i] * stack_limit,
                                                         [models_data['MODEL'][i]] * items_per_stack, overhang))

            # We save the number of individual crates to build and convert it into
            # integer to avoid conflict with range function
            nbr_individual_crates = int((qty - (items_per_stack * nbr_stacks)) / nbr_per_crate)

            for j in range(nbr_individual_crates):

                # We build the crate and send it to the crates manager
                remaining_crates.add_crate(LoadingObjects.Crate([models_data['MODEL'][i]] * nbr_per_crate,
                                                                max(models_data['LENGTH'][i], models_data['WIDTH'][i]),
                                                                min(models_data['WIDTH'][i], models_data['LENGTH'][i]),
                                                                models_data['HEIGHT'][i], stack_limit, overhang))

    return model_names, warehouse, remaining_crates


def read_trailers_data(ws, oh_authorized=72, max_length=636):

    """
    This function read trailers' data from worksheet

    :param ws: Active Excel worksheet that contains data
    :param oh_authorized: Maximal measure of overhanging authorized by law (in inches)
    :param max_length: Maximal length of a load including overhang (in inches)
    :return: list of trailers

    """

    # We initialize a list of trailers
    trailers = []

    # We access to the table where the data on trailers is contained (table 1)
    trailers_range = ws._tables[1].ref

    # We save the data in a data frame
    trailers_data = build_dataframe(ws[trailers_range])

    # For every lines of the data frame
    for i in trailers_data.index:

        # We save the quantity
        qty = trailers_data['QTY'][i]

        if qty > 0:

            # We save trailer's length
            t_length = trailers_data['LENGTH'][i]

            # We compute overhanging measure allowed for the trailer
            trailer_oh = min(max_length - t_length, oh_authorized)

            # We build "qty" trailer that we add to the trailers list
            for j in range(0, qty):
                trailers.append(LoadingObjects.Trailer(trailers_data['CATEGORY'][i], t_length,
                                                       trailers_data['WIDTH'][i], trailers_data['HEIGHT'][i],
                                                       trailers_data['PRIORITY_RANK'][i], trailer_oh))
    return trailers


def read_max_nbr_of_trailer(ws):

    """
    Save the maximal number of trailer to produce
    :param ws: Active Excel worksheet that contains data
    :return: maximal number of trailer to produce (int)
    """

    # We access to the table where the data on maximal number of trailer is contained (table 3)
    max_nbr_range = ws._tables[3].ref

    # We save the data contained in the table
    max_data = build_dataframe(ws[max_nbr_range])
    max_nbr = max_data['MAX TRAILERS'][0]

    # If the table is empty, we consider that there is no maximum and return a very high value
    if not max_nbr:
        return 5000

    # If the value is positive and is an integer
    elif isinstance(max_nbr, numbers.Number) and max_nbr > 0:
        return int(max_nbr)

    # Else we raise an error message
    else:
        raise Exception('The maximal number of trailer specified is incorrect')


def write_summarized_data(Model_names, Trailers, Unused_list, ws, directory):

    """
    Writes a loading summary in a .xlsx file and create a folder named "P2P - <date>"
    at the directory mentioned and save the file in it.

    :param Trailers: List of trailers
    :param Unused_list:  List of models unused
    :param ws:  Active Excel worksheet that contains data
    :param directory: String that mention path where the created folder is going to be saved

    """

    # We initialize a data frame with column names needed
    data_frame = pd.DataFrame(columns=(["TRAILER", "TRAILER LENGTH", "LOAD LENGTH"]+Model_names))

    # We initialize an indec
    i = 0

    # We add a line in the dataframe for every trailer used
    for trailer in Trailers:

        # We save the quantities of every models inside the trailer
        s = trailer.load_summary()

        # Every line of dataframe has the category of trailer, his length, his remaining_length (in feets)
        # and the quantities of every models in it.
        data_frame.loc[i] = [trailer.category] + [round(trailer.length/12, 1)] + \
                            [round(trailer.length_used/12, 1)] + \
                            [s[model] if s[model] > 0 else '' for model in Model_names]
        i += 1

    # We add a line for unused model
    unused_summary = Counter(Unused_list)
    data_frame.loc[i] = ["REMAINING", '', ''] + \
                        [unused_summary[model] if unused_summary[model] > 0 else '' for model in Model_names]

    # We execute a groupby with trailer in the same category
    data_frame = data_frame.groupby(data_frame.columns.tolist()).size().to_frame('QTY').reset_index()

    # We rearrange columns
    cols = data_frame.columns.tolist()
    cols = cols[0:1] + cols[-1:] + cols[1:-1]
    data_frame = data_frame[cols]

    # We set indexes
    data_frame.set_index("TRAILER", inplace=True)

    # We erase the quantity of the QTY column
    data_frame.loc[["REMAINING"], ['QTY']] = ''

    # We save the names of "plant to plant" and date from the specific table in worksheet (table 2)
    title_infos_range = ws._tables[2].ref
    title_infos = build_dataframe(ws[title_infos_range])
    plan_A = title_infos['FROM'][0]
    plan_B = title_infos['TO'][0]
    file_date = str(title_infos['DATE'][0])
    file_date = file_date[0:10]                   # Date without time

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
    title = path + "P2P from " + plan_A + " to " + plan_B + " " + file_date + ".xlsx"

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


def save_remaining(unused_list, nbr_trailers, execution_time):

    """
    Saves the remaining models in a text file called "remaining_models.txt"

    :param unused_list: List of models unused
    :param nbr_trailers: Number of trailers used (int)
    :param execution_time: Time of execution (in seconds)

    """

    # We open the file text (or create it if it doesn't exist)
    f = open("remaining_models.txt", "w+")

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

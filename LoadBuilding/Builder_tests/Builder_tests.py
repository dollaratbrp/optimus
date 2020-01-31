"""

Created by Nicolas Raymond on 2019-11-22.

This file provides tests to validate LoadBuilder functions.

Last update : 2019-11-22
By : Nicolas Raymond

"""
import sys
import os
from LoadBuilder import LoadBuilder
from openpyxl import load_workbook
from copy import deepcopy

# Import read_and_write
wd = os.path.dirname(sys.argv[0])
sys.path.append(wd)
import Read_And_Write as rw


def main():

    # We get data on models and trailers
    wb = load_workbook(wd + '\\Models_and_trailers.xlsx', data_only=True)
    ws = wb.active
    models_data = rw.read_models_data(ws)
    trailers_data = rw.read_trailers_data(ws)

    # We look the data frame to see if they are correct
    # rw.display_df(trailers_data)

    # We initialize a dummy load builder to manage plant to plant between Juarez and El Paso
    lb1 = LoadBuilder(trailers_data)

    # We build loads
    res = lb1.build(models_data, 2, plot_load_done=True)
    print(lb1.get_loading_summary(), '\n')

    for trailer in lb1.trailers_done:
        print([stack.models for stack in trailer.load])

    models_data.loc[0, 'QTY'] = 1

    print('\n', models_data, '\n')

    lb1.patching_activated = True
    res = lb1.build(models_data, 0)

    for trailer in lb1.trailers_done:
        trailer.plot_load()
        print([stack.models for stack in trailer.load])

    # We look at the loading summaries
    lb1_summary = lb1.get_loading_summary()
    print('\n', lb1_summary)


if __name__ == "__main__":
    main()


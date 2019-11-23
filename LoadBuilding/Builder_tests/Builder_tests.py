"""

Created by Nicolas Raymond on 2019-11-22.

This file provides tests to validate LoadBuilder functions.

Last update : 2019-11-22
By : Nicolas Raymond

"""
import sys
import os
import time
from LoadBuilder import LoadBuilder
from openpyxl import load_workbook

# Import read_and_write
wd = os.getcwd()
sys.path.append(wd)
import Read_And_Write as rw


def main():

    # We get data on models and trailers
    workbook_path = os.path.join(wd, "Models_and_trailers.xlsx")
    wb = load_workbook(workbook_path, data_only=True)
    ws = wb.active
    models_data = rw.read_models_data(ws)
    trailers_data = rw.read_trailers_data(ws)

    # We look the data frame to see if they are correct
    #rw.display_df(models_data)
    #rw.display_df(trailers_data)

    # We initialize a dummy load builder to manage plant to plant between Juarez and El Paso
    lb1 = LoadBuilder('Juarez', 'El Paso', models_data, trailers_data, '2019-11-23')

    # We build loads
    res, time = lb1.build(10, 3, plot_load_done=True)
    rw.show_summary(res, len(lb1), time)


    # We look again at the data frames to see if the were correctly updated
    #rw.display_df(models_data)
    #rw.display_df(trailers_data)

    # We initialize a dummy load builder to manage plant to plant between Juarez and Juarez 2
    lb2 = LoadBuilder('Juarez', 'Juarez 2', models_data, trailers_data, '2019-11-23')

    # We build loads
    res, time = lb2.build(10, 3, plot_load_done=True)
    rw.show_summary(res, len(lb2), time)

    # We look a last time at the data frames to see if the were correctly updated
    #rw.display_df(models_data)
    #rw.display_df(trailers_data)

    # We save the results of loading
    res_directory = os.path.join(wd, "Results", '')
    lb1.write_summarized_data(res_directory)
    lb2.write_summarized_data(res_directory)


if __name__ == "__main__":
    main()


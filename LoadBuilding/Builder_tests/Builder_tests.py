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
    # rw.display_df(trailers_data)

    # We initialize a dummy load builder to manage plant to plant between Juarez and El Paso
    lb1 = LoadBuilder('Juarez', 'El Paso', trailers_data)

    # We build loads
    res = lb1.build(models_data, 10, plot_load_done=False)
    print(res)

    # We look again at the data frames to see if the were correctly updated
    # rw.display_df(trailers_data)

    # We initialize a dummy load builder to manage plant to plant between Juarez and Juarez 2
    lb2 = LoadBuilder('Juarez', 'Juarez 2', trailers_data)

    # We build loads
    res = lb2.build(models_data, 10, plot_load_done=False)
    print(res)

    # We look a last time at the data frames to see if the were correctly updated
    # rw.display_df(trailers_data)

    # We look at the loading summaries
    lb1_summary = lb1.get_loading_summary()
    lb2_summary = lb2.get_loading_summary()
    rw.display_df(lb1_summary)
    rw.display_df(lb2_summary)

    # We save the results of loading
    res_directory = os.path.join(wd, "Results", '')
    rw.write_summarized_data(lb1_summary, lb1.plant_from, lb1.plant_to, res_directory)
    rw.write_summarized_data(lb2_summary, lb2.plant_from, lb2.plant_to, res_directory)


if __name__ == "__main__":
    main()


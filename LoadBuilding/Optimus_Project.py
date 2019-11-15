import os
import time
import LoadingProcess as Loading
import Read_And_Write as rw
from openpyxl import load_workbook

"""

Project Optimus

Created by Nicolas Raymond on 2019-05-31.

The goal is to maximize the volume of items shipped while minimizing number of trailers used.

It is strongly recommend to read LoadingObjects.py, Loading.py and Read_And_Write.py before.

Last update : 2019-11-01
By : Nicolas Raymond

"""


def main():
    
    # Initialization of global variables
    workbook_path = 'C:\\Users\\raymoni.BRP\\Desktop\\loading_optimize\\P2P_management_v02.xlsm'
    archive_folder_path = 'C:\\Users\\raymoni.BRP\\Desktop\\loading_optimize\\email_archive\\'

    max_overhanging_measure = 72  # Maximal overhanging measure authorized (in inches) for a trailer
    max_trailer_length = 636  # Maximal length of a trailer (in inches)
    plc_lb = 0.75  # Lowest percentage of length that must be covered (lb = 'lower bound')

    # \\-------------------------------------------------------------------------------------------------------------//

    # We save start time
    start_time = time.time()

    # We load the workbook where the data for loading is contained
    wb = load_workbook(workbook_path, data_only=True)

    # We load the worksheet that is open at the moment
    ws = wb.active

    # We initialize a list that will contain all the names of the models unused in the loading process
    unused = []

    # We read data about trailers and models in the worksheet
    model_names, warehouse, remaining_crates = rw.read_models_data(ws)
    trailers = rw.read_trailers_data(ws, max_overhanging_measure, max_trailer_length)
    max_trailers = rw.read_max_nbr_of_trailer(ws)
   
    # We execute the loading process
    Loading.solve(warehouse, trailers, remaining_crates, unused, plc_lb)

    # If a maximal number of trailer N is specified in the work sheet we choose
    # the N best loadings according to the number of units
    Loading.select_top_n(trailers, unused, max_trailers)

    # We save the results in a .xlsx file
    rw.write_summarized_data(model_names, trailers, unused, ws, archive_folder_path)

    # We save the end of execution time
    end_time = time.time()

    # We save the number of trailer used and the time of execution
    execution_time = end_time - start_time
    nbr_trailers = len(trailers)

    # We print a summary of the results in a .txt file
    rw.save_remaining(unused, nbr_trailers, execution_time)

    # We open the .txt file
    os.startfile("remaining_models.txt")


if __name__ == "__main__":
    main()





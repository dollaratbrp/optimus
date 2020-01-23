"""

Author : Olivier Lefebre

This file creates loads for P2P

Last update : 2020-01-06
By : Nicolas Raymond

"""

import sys
from ParametersBox import OpenParameters, MissingP2PBox
from P2PFunctions import *
from ProcessValidation import validate_process
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl import Workbook
import os

""" Used SQL Table and actions: """
# [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS] : Get data, alter, delete, send
# [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS] : Get data, alter, delete, send
# [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY] : Get data
# [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY] : Get data
# OTD_1_P2P_D_INCLUDED_INVENTORY : Get data
# OTD_1_P2P_F_HISTORICAL : Send data

# Folder where the excel workbook is saved
saveFolder = 'S:\Shared\Business_Planning\Personal\Raymond\P2P\\'

# Parameters variables
dayTodayComplete = pd.datetime.now().replace(second=0, microsecond=0)  # date to set in SQL for import date
dayToday = weekdays(0)  # Date to display in report
printLoads = False  # Print created loads
AutomaticRun = False  # set to True to automate code
validation = True     # set to True to validate the results received after the process
dest_filename = 'P2P_Summary_'+dayToday  # Name of excel file with today's date


def p2p_full_process():
    """
    Executes P2P full process

    :return: summary of the full process in at the 'saveFolder' directory

    """
    if not AutomaticRun:
        if not OpenParameters():  # If user cancel request
            sys.exit()

    timeSinceLastCall('', False)

    ####################################################################################################################
    #                                                Get SQL DATA
    ####################################################################################################################

    # If SQL queries crash
    downloaded = False
    nbr_of_try = 0
    while not downloaded and nbr_of_try < 3:  # 3 trials, SQL Queries sometime crash for no reason
        nbr_of_try += 1
        try:
            downloaded = True

            ####################################################################################
            #                     Email address Query
            ####################################################################################

            email_connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning',
                                             'OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS', headers='EMAIL_ADDRESS')

            email_query = """ SELECT distinct [EMAIL_ADDRESS]
             FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS]
             WHERE PROJECT = 'P2P'
            """

            # GET SQL DATA
            email_data = email_connection.GetSQLData(email_query)
            emails_list = [item for sublist in email_data for item in sublist]

            ####################################################################################
            #                     Parameters Query
            ####################################################################################

            DATAParams, param_connection = get_parameter_grid()

            ####################################################################################
            #                     Parameters P2P ORDER (for excel sheet order)
            ####################################################################################

            p2p_order_query = """ SELECT distinct  [POINT_FROM],[POINT_TO]
              FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
              where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
              and SKIP = 0
              order by [POINT_FROM],[POINT_TO]
            """
            # GET SQL DATA
            P2POrder = param_connection.GetSQLData(p2p_order_query)

            ####################################################################################
            #                     WishList recuperation
            ####################################################################################

            DATAWishList = get_wish_list()

            ####################################################################################
            #                     Inventory recuperation
            ####################################################################################

            DATAINV = get_inventory_and_qa()

            ####################################################################################
            #                     Nested Shipping_point recuperation
            ####################################################################################

            global DATAInclude
            get_nested_source_points(DATAInclude)

            ####################################################################################
            #  Look if all point_from + shipping_point are in parameters
            ####################################################################################

            SQLMissing = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PRIORITY', headers='')
            QueryMissing = """SELECT DISTINCT [POINT_FROM]
                  ,[SHIPPING_POINT]
              FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY] 
              where CONCAT(POINT_FROM,SHIPPING_POINT) not in (
                select distinct CONCAT( [POINT_FROM],[POINT_TO])
                FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS] where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]) )
                and [POINT_FROM] <>[SHIPPING_POINT] 
            """
            DATAMissing = SQLMissing.GetSQLData(QueryMissing)

        except:
            downloaded = False
            print('SQL Query failed')

    # If SQL Queries failed
    if not downloaded:
        try:
            send_email(emails_list, dest_filename, 'SQL QUERIES FAILED')
        except:
            pass
        sys.exit()

    timeSinceLastCall('Get SQL DATA')

    # If there are missing P2P in parameters table
    if not AutomaticRun:
        if DATAMissing:
            if not MissingP2PBox(DATAMissing):
                sys.exit()

    timeSinceLastCall('', False)

    ####################################################################################################################
    #                                                 Excel Workbook declaration
    ####################################################################################################################

    # Initialization of workbook and filling option to warn user in output
    wb = Workbook()
    Warning_fill = PatternFill(fill_type="solid", start_color="FFFF00", end_color="FFFF00")

    # Summary
    summary_ws = wb.active
    summary_ws.title = "SUMMARY"
    worksheet_formatting(summary_ws, ['POINT_FROM', 'SHIPPING_POINT', 'NUMBER_OF_LOADS'], [13, 16, 19])

    # Approved loads
    approved_ws = wb.create_sheet("APPROVED")
    columns_title = ['POINT_FROM', 'SHIPPING_POINT', 'LOAD_NUMBER', 'CATEGORY', 'LOAD_LENGTH',
                     'MATERIAL_NUMBER', 'QUANTITY', 'SIZE_DIMENSIONS',
                     'SALES_DOCUMENT_NUMBER', 'SALES_ITEM_NUMBER', 'SOLD_TO_NUMBER']
    columns_width = [20]*(len(columns_title)-1) + [27]
    worksheet_formatting(approved_ws, columns_title, columns_width)

    # Unbooked
    unbooked_ws = wb.create_sheet("UNBOOKED")
    worksheet_formatting(unbooked_ws, ['POINT_FROM', 'MATERIAL_NUMBER', 'QUANTITY'], [20, 20, 12])

    # Booked unused
    unused_ws = wb.create_sheet("BOOKED_UNUSED")
    worksheet_formatting(unused_ws, ['POINT_FROM', 'MATERIAL_NUMBER', 'QUANTITY'], [20, 20, 12])

    ####################################################################################################################
    #                                            Isolate perfect match
    ####################################################################################################################

    # We initialize a list that will contain all wish approved
    ListApprovedWish = find_perfect_match(DATAWishList, DATAINV, DATAParams)

    ####################################################################################################################
    #                                                Create Loads
    ####################################################################################################################

    for param in DATAParams:  # for all P2P in parameters

        # Initialization of empty list
        temporary_on_load = []  # List to remember the INVobjs that will be sent to the LoadBuilder
        loadbuilder_input = []  # List that will contain the data to build the frame we'll send to the LoadBuilder

        # Initialization of an empty ranking dictionary
        ranking = {}

        # We loop through our wishes list
        for wish in ListApprovedWish:

            # If the wish is not fulfilled and his POINT FROM and POINT TO are corresponding with the param (p2p)
            if wish.QUANTITY > 0 and wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO:
                temporary_on_load.append(wish)

                # Here we set QTY and NBR_PER_CRATE to 1 because each line of the wishlist correspond to
                # one crate and not one unit! Must be done this way to avoid having getting to many size_code
                # in the returning list of the LoadBuilder
                loadbuilder_input.append(wish.get_loadbuilder_input_line())

                # We add the ranking of the wish in the ranking dictionary
                if wish.SIZE_DIMENSIONS in ranking:
                    ranking[wish.SIZE_DIMENSIONS] += [wish.RANK]
                else:
                    ranking[wish.SIZE_DIMENSIONS] = [wish.RANK]

        # Construction of the data frame which we'll send to the LoadBuilder of our parameters object (p2p)
        input_dataframe = loadbuilder_input_dataframe(loadbuilder_input)

        # Create loads
        result = param.LoadBuilder.build(input_dataframe, param.LOADMAX, ranking=ranking, plot_load_done=printLoads)

        # Choose which wish to send in load based on selected crates and priority order
        for model in result:
            found = False
            for OnLoad in temporary_on_load:
                if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0:
                    OnLoad.QUANTITY = 0
                    found = True
                    param.AssignedWish.append(OnLoad)
                    break
            if not found:  # If one returned crate doesn't match data
                print('Error in Perfect Match: impossible result.\n')

    ####################################################################################################################
    #                                 Store unallocated units in inv pool
    ####################################################################################################################
    # If a wish item is not on a load, we give back his reserved inv
    for wish in ListApprovedWish:
        if wish.QUANTITY > 0:
            for inv in wish.INV_ITEMS:
                inv.QUANTITY += 1
            wish.INV_ITEMS = []

    ####################################################################################################################
    #                             Try to Make the minimum number of loads for each P2P
    ####################################################################################################################

    satisfy_max_or_min(DATAWishList, DATAINV, DATAParams, print_loads=printLoads)

    ####################################################################################################################
    #                             Try to Make the maximum number of loads for each P2P
    ####################################################################################################################

    satisfy_max_or_min(DATAWishList, DATAINV, DATAParams, satisfy_min=False, print_loads=printLoads)

    ####################################################################################################################
    #                                           Writing of the results
    ####################################################################################################################

    # We display loads create in each p2p for our own purpose
    print('\n\nResults')
    for param in DATAParams:
        print('\n\n')
        print(param.POINT_FROM, ' _ ', param.POINT_TO)
        print(len(param.LoadBuilder))
        print(param.LoadBuilder.trailers_done)
        print(param.LoadBuilder.get_loading_summary())

    lineIndex = 2  # To display a warning if number of loads is lower than parameters min


    # SQL to send DATA
    headersResult = 'POINT_FROM,SHIPPING_POINT,LOAD_NUMBER,MATERIAL_NUMBER,QUANTITY,SIZE_DIMENSIONS,' \
                    'SALES_DOCUMENT_NUMBER,SALES_ITEM_NUMBER,SOLD_TO_NUMBER,IMPORT_DATE'

    SQLResult = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_HISTORICAL', headers=headersResult)

    for order in P2POrder:  # To send data in excel workbook, order by : -point_from , -point_to
        for param in DATAParams:
            if param.POINT_FROM == order[0] and param.POINT_TO == order[1]:

                LoadIteration = 0  # Number of each load (reset for each plant to plant)

                # section for summary worksheet
                summary_ws.append([param.POINT_FROM, param.POINT_TO, len(param.LoadBuilder)])
                if len(param.LoadBuilder) < param.LOADMIN:
                    summary_ws.cell(row=lineIndex, column=3).fill = Warning_fill
                lineIndex += 1
                # Approved worksheet
                loads = param.LoadBuilder.get_loading_summary()
                if len(param.LoadBuilder) > 0:  # If some loads were created
                    for line in range(len(loads)):
                        for Iteration in range(int(loads["QTY"][line])):  # For every load of this kind
                            LoadIteration += 1
                            for column in loads.columns[4::]:  # loop through size_dimensions
                                if loads[column][line] != '':  # if there is some quantity
                                    for QUANTITY in range(int(loads[column][line])):  # For every unit of this crate on the load
                                        for wish in param.AssignedWish:  # we associate a wish unit to this crate, then we save it
                                            if not wish.Finished and wish.SIZE_DIMENSIONS == column:
                                                wish.Finished = True
                                                valuesSQL = [(param.POINT_FROM, param.POINT_TO, LoadIteration,
                                                              wish.MATERIAL_NUMBER, wish.ORIGINAL_QUANTITY,
                                                              column, wish.SALES_DOCUMENT_NUMBER,
                                                              wish.SALES_ITEM_NUMBER,
                                                              wish.SOLD_TO_NUMBER, dayTodayComplete)]

                                                approved_ws.append([param.POINT_FROM, param.POINT_TO, LoadIteration,
                                                                   loads['TRAILER'][line], loads['LOAD LENGTH'][line],
                                                                   wish.MATERIAL_NUMBER, wish.ORIGINAL_QUANTITY,
                                                                   column, wish.SALES_DOCUMENT_NUMBER,
                                                                   wish.SALES_ITEM_NUMBER, wish.SOLD_TO_NUMBER])

                                                SQLResult.sendToSQL(valuesSQL)
                                                break

                break

    # Assign left inv to wishList, to look for booked but unused
    for wish in DATAWishList:
        if wish.QUANTITY == 0 and not wish.Finished:
            print('Error with wish: ', wish.lineToXlsx())
        if wish.QUANTITY > 0:
            position = 0
            for Iteration in range(wish.QUANTITY):
                for It, inv in enumerate(DATAINV[position::]):
                    if EquivalentPlantFrom(inv.POINT, wish.POINT_FROM) and \
                            wish.MATERIAL_NUMBER == inv.MATERIAL_NUMBER and inv.QUANTITY - inv.unused > 0:
                        if inv.Future:  # QA of tomorrow, need to look if load is for today or later
                            InvToTake = False
                            for param in DATAParams:
                                if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO\
                                        and param.days_to > 0:
                                    InvToTake = True
                                    break
                            if InvToTake:
                                inv.unused += 1
                                position += It
                                break  # no need to look further
                        else:
                            inv.unused += 1
                            position += It
                            break  # no need to look further

    # send inv in unused_ws and unbooked_ws
    for inv in DATAINV:
        if inv.unused > 0:
            unused_ws.append([inv.POINT, inv.MATERIAL_NUMBER, inv.unused])
        if inv.QUANTITY - inv.unused > 0:
            unbooked_ws.append([inv.POINT, inv.MATERIAL_NUMBER, inv.QUANTITY-inv.unused])

    # We save the workbook and the reference
    reference = [savexlsxFile(wb, saveFolder, dest_filename)]

    # We send the emails
    send_email(emails_list, dest_filename, '', reference)

    # We validate the process' results if the user wants to
    if validation:
        workbook_path = saveFolder + dest_filename + '.xlsx'
        validate_process(workbook_path)

    # We open excel workbook
    os.system('start "excel" "'+str(reference[0])+'"')

"""
This file manages all activities linked to Forecast

Author : Olivier Lefebvre

Update by : Nicolas Raymond
"""

# P2P Forecast for next 8 weeks

#  Used SQL Table and actions:
# [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS] : Get data, alter, delete, send
# [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS] : Get data, alter, delete, send
# [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY] : Get data
# [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY] : Get data
# OTD_1_P2P_D_INCLUDED_INVENTORY : Get data
# OTD_1_P2P_F_FORECAST_LOADS : Send data
# OTD_1_P2P_F_FORECAST_PRIORITY : Send data

# Path where the forecast results are saved
saveFolder = 'S:\Shared\Business_Planning\Personal\Raymond\P2P\\'

# from openpyxl.utils.dataframe import dataframe_to_rows
from ParametersBox import *
from P2PFunctions import *
from openpyxl import Workbook, load_workbook
from tqdm import tqdm
import numpy as np

AutomaticRun = False  # set to True to automate code


def forecast():
    """
    Runs the forecast for the next 8 weeks

    """
    if not AutomaticRun:
        if not OpenParameters('FORECAST'):  # If user cancel request
            sys.exit()
        IsAdhoc = 1  # 1 for True, 0 for False (if automatic)
    else:
        IsAdhoc = 0  # 1 for True, 0 for False (if automatic)

    # -----------------------------------------------------------------------------------------------------------------#
    # -----------------------------------------------------------------------------------------------------------------#
    #                                             MODIFICATIONS SECTION
    # -----------------------------------------------------------------------------------------------------------------#
    # -----------------------------------------------------------------------------------------------------------------#

    dayTodayComplete = pd.datetime.now().replace(second=0, microsecond=0)
    dayToday = weekdays(0)
    printLoads = False

    dest_filename = 'P2P_Forecast_'+dayToday  # Email subject with today's date

    # -----------------------------------------------------------------------------------------------------------------#
    # -----------------------------------------------------------------------------------------------------------------#
    #                                            END OF MODIFICATIONS
    # -----------------------------------------------------------------------------------------------------------------#
    # -----------------------------------------------------------------------------------------------------------------#

    GeneralErrors = ''  # General errors met during execution

    ####################################################################################################################
    #                                                Get SQL DATA
    ####################################################################################################################

    # #If SQL queries crash
    downloaded = False
    numberOfTry = 0
    while not downloaded and numberOfTry < 3:  # 3 trials, SQL Queries sometime crash for no reason
        numberOfTry += 1
        try:
            downloaded = True

            ####################################################################################
            #                                 Email address Query
            ####################################################################################

            EmailList = get_emails_list('FORECAST')

            ####################################################################################
            #                                     Parameters Query
            ####################################################################################

            headerParams = 'POINT_FROM,POINT_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,TRANSIT,PRIORITY_ORDER,SKIP'
            SQLParams = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS',
                                      headers=headerParams)

            QueryParams = """ SELECT  [POINT_FROM]
                  ,[POINT_TO]
                  ,[LOAD_MIN]
                  ,[LOAD_MAX]
                  ,[DRYBOX]
                  ,[FLATBED]
                  ,[TRANSIT]
                  ,[PRIORITY_ORDER]
                  ,DAYS_TO
              FROM [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS]
              where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS])
              and SKIP = 0
              order by PRIORITY_ORDER
            """
            # GET SQL DATA
            DATAParams = [Parameters(*sublist) for sublist in SQLParams.GetSQLData(QueryParams)]

            ####################################################################################
            #                               Distinct date
            ####################################################################################

            #  We use the dates to loop through time
            SQLINV = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_INVENTORY',
                                      headers='')
            QueryDATE = """ SELECT distinct convert(date,[AVAILABLE_DATE]) as [AVAILABLE_DATE]
                      FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
                      where [AVAILABLE_DATE] between convert(date,getdate()+1) and convert(date,getdate() + 7*8)
                      order by [AVAILABLE_DATE]
                    """

            DATADATE = [single for sublist in SQLINV.GetSQLData(QueryDATE) for single in sublist]

            ####################################################################################
            #                                First Inv Query
            ####################################################################################
            QueryInv = """ select distinct SHIPPING_POINT
                  ,[MATERIAL_NUMBER]
                  ,case when sum(tempo.[QUANTITY]) <0 then 0 else convert(int,sum(tempo.QUANTITY)) end as [QUANTITY]
                  , [AVAILABLE_DATE]
                  ,[STATUS]
                  from(
              SELECT  [SHIPPING_POINT]
                  ,[MATERIAL_NUMBER]
                  , [QUANTITY]
                  ,convert(DATE,[AVAILABLE_DATE]) as [AVAILABLE_DATE]
                  ,[STATUS]
              FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
              where status = 'INVENTORY'
             UNION(
              SELECT distinct [SHIPPING_POINT]
                  ,[MATERIAL_NUMBER]
                  ,0 as [QUANTITY]
                  ,convert(date,getdate()) as [AVAILABLE_DATE]
                  ,'INVENTORY' as [STATUS]
              FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
              )) as tempo
             group by SHIPPING_POINT,[MATERIAL_NUMBER],  [AVAILABLE_DATE], [STATUS]
             order by SHIPPING_POINT, MATERIAL_NUMBER
                            """

            DATAINV = [INVObj(*obj) for obj in SQLINV.GetSQLData(QueryInv)]

            ####################################################################################
            #                                Prod & QA Query
            ####################################################################################
            # Filter based on if production is late, to make later
            QueryProd = """ SELECT  [SHIPPING_POINT]
                                  ,[MATERIAL_NUMBER]
                                  ,[QUANTITY]
                                  ,convert(date,[AVAILABLE_DATE]) as AVAILABLE_DATE
                                  ,[STATUS]
                              FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
                              where status in ('QA HOLD','PRODUCTION PLAN')
                              and AVAILABLE_DATE between convert(date,getdate()) and convert(date, getdate() + 9*7)
                              order by AVAILABLE_DATE,SHIPPING_POINT,MATERIAL_NUMBER
                                    """

            DATAProd = [INVObj(*obj) for obj in SQLINV.GetSQLData(QueryProd)]

            ####################################################################################
            #                                 WishList Query
            ####################################################################################

            DATAWishList = get_wish_list(forecast=True)

            ####################################################################################
            #                             Included Shipping_point
            ####################################################################################

            global DATAInclude
            get_nested_source_points(DATAInclude)

        except:
            downloaded = False
            print('SQL Query failed')

    # If SQL Queries failed
    if not downloaded:
        try:
            print('failed')
            send_email(EmailList, dest_filename, 'SQL QUERIES FAILED')
        except:
            pass
        sys.exit()

    ####################################################################################################################
    #                                     SQL tables declaration and cleaning
    ####################################################################################################################
    headerLoads = 'POINT_FROM,SHIPPING_POINT,QUANTITY,DATE,IMPORT_DATE,IS_ADHOC'
    SQLLoads = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_FORECAST_LOADS', headers=headerLoads)
    SQLLoads.deleteFromSQL()

    headerPRIORITY = 'SALES_DOCUMENT_NUMBER,SALES_ITEM_NUMBER,SOLD_TO_NUMBER,POINT_FROM,SHIPPING_POINT,DIVISION,'\
                     'MATERIAL_NUMBER,SIZE_DIMENSIONS,LENGTH,WIDTH,HEIGHT,STACKABILITY,OVERHANG,QUANTITY,' \
                     'PRIORITY_RANK,X_IF_MANDATORY,PROJECTED_DATE,IMPORT_DATE,IS_ADHOC'
    SQLPRIORITY = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_FORECAST_PRIORITY',
                                headers=headerPRIORITY)
    SQLPRIORITY.deleteFromSQL()

    ####################################################################################################################
    #                                           Excel Workbook declaration
    ####################################################################################################################

    # Initialization of workbook and filling option to warn user in output
    wb = Workbook()

    # Summary
    summary_ws = wb.active
    summary_ws.title = "SUMMARY"
    summary_columns = ['POINT_FROM', 'SHIPPING_POINT', 'QUANTITY', 'DEPARTURE_DATE']

    # Detailed
    detailed_ws = wb.create_sheet("DETAILED")
    detailed_columns = ['POINT_FROM', 'SHIPPING_POINT', 'DIVISION', 'MATERIAL_NUMBER',
                        'SIZE_DIMENSIONS', 'QUANTITY', 'DEPARTURE_DATE']

    columns_width = [20]*(len(detailed_columns)-1) + [25]
    worksheet_formatting(detailed_ws, detailed_columns, columns_width)

    ####################################################################################################################
    #                                                Main loop for everyday
    ####################################################################################################################

    # We add today's prod and QA to inventory, but we don't create load for today (we make them for tomorrow or later)
    for prod in DATAProd:
        if prod.DATE > weekdays(0):  # Ordered by date, so no need to continue
            break
        elif prod.DATE == weekdays(0) and prod.QUANTITY > 0 and prod.STATUS == 'QA HOLD':
            for inv in DATAINV:
                if inv.POINT == prod.POINT and inv.MATERIAL_NUMBER == prod.MATERIAL_NUMBER:
                    inv.QUANTITY += prod.QUANTITY
                    prod.QUANTITY = 0
                    break  # we found the good inv

    # Inventory is now updated.

    # We initialize a list of data that will be used to create the pivot table with pandas and a column counter
    summary_data = []
    column_counter = 3  # Starts at 3 because of the columns point from and point to

    # Initialization of the load bar
    progress = tqdm(total=len(DATADATE), desc='Forecast progress')
    for date in DATADATE:

        # We reset the number of shared flatbed_53
        reset_flatbed_53()

        # we add today's prod and QA to inventory
        for prod in DATAProd:
            if prod.DATE > date:  # Ordered by date, so no need to continue
                break
            elif prod.DATE == date and prod.QUANTITY > 0:
                for inv in DATAINV:
                    if inv.POINT == prod.POINT and inv.MATERIAL_NUMBER == prod.MATERIAL_NUMBER:
                        inv.QUANTITY += prod.QUANTITY
                        prod.QUANTITY = 0
                        break  # we found the good inv

        # Inventory is now updated.

        # Now we create new loadBuilders
        for param in DATAParams:
            param.new_LoadBuilder()

        # Isolate perfect match
        ListApprovedWish = []

        for wish in DATAWishList:
            if wish.QUANTITY > 0:
                position = 0  # to not loop again through the complete list
                for Iteration in range(wish.QUANTITY):
                    for It, inv in enumerate(DATAINV[position::]):
                        if EquivalentPlantFrom(inv.POINT, wish.POINT_FROM) and \
                                wish.MATERIAL_NUMBER == inv.MATERIAL_NUMBER and inv.QUANTITY > 0:
                            inv.QUANTITY -= 1
                            wish.INV_ITEMS.append(inv)
                            position += It
                            break  # no need to look further

                if len(wish.INV_ITEMS) < wish.QUANTITY:  # We give back taken inv
                    for invToGiveBack in wish.INV_ITEMS:
                        invToGiveBack.QUANTITY += 1
                    wish.INV_ITEMS = []
                else:
                    ListApprovedWish.append(wish)

        # We now create loads with those perfect match
        for param in DATAParams:
            tempoOnLoad = []
            invData = []

            for wish in ListApprovedWish:
                if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO and wish.QUANTITY > 0:
                    tempoOnLoad.append(wish)
                    invData.append(wish.get_loadbuilder_input_line())  # quantity is one, one box for each line
            models_data = loadbuilder_input_dataframe(invData)
            result = param.LoadBuilder.build(models_data, param.LOADMAX, plot_load_done=printLoads)

            for model, crate_type in result:
                found = False
                for OnLoad in tempoOnLoad:
                    # we send allocated wish units to sql
                    if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0 and OnLoad.CRATE_TYPE == crate_type:
                        OnLoad.QUANTITY = 0
                        found = True
                        param.AssignedWish.append(OnLoad)
                        OnLoad.EndDate = weekdays(max(0, param.days_to-1), officialDay=date)
                        SQLPRIORITY.sendToSQL(OnLoad.lineToXlsx(dayTodayComplete))
                        detailed_ws.append(OnLoad.lineToXlsx(dayTodayComplete, filtered=True))
                        DATAWishList.remove(OnLoad)
                        break
                if not found:
                    GeneralErrors += 'Error in Perfect Match: impossible result.\n'

        # Store unallocated units in inv pool
        for wish in ListApprovedWish:
            if wish.QUANTITY > 0:
                for inv in wish.INV_ITEMS:
                    inv.QUANTITY += 1
                wish.INV_ITEMS = []

        # Try to Make the minimum number of loads for each P2P
        for param in DATAParams:
            if len(param.LoadBuilder) < param.LOADMIN:  # If the minimum isn't reached
                tempoOnLoad = []
                invData = []

                # create data table
                for wish in DATAWishList:
                    if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO and wish.QUANTITY > 0:
                        position = 0
                        for Iteration in range(wish.QUANTITY):
                            for It, inv in enumerate(DATAINV[position::]):
                                if EquivalentPlantFrom(inv.POINT, wish.POINT_FROM) and \
                                        inv.MATERIAL_NUMBER == wish.MATERIAL_NUMBER and inv.QUANTITY > 0:
                                    inv.QUANTITY -= 1
                                    wish.INV_ITEMS.append(inv)
                                    position += It
                                    break  # no need to look further
                        if len(wish.INV_ITEMS) < wish.QUANTITY:  # We give back taken inv
                            for invToGiveBack in wish.INV_ITEMS:
                                invToGiveBack.QUANTITY += 1
                            wish.INV_ITEMS = []
                        else:
                            tempoOnLoad.append(wish)
                            invData.append(
                                wish.get_loadbuilder_input_line())
                # Create loads
                models_data = loadbuilder_input_dataframe(invData)
                result = param.LoadBuilder.build(models_data, param.LOADMIN - len(param.LoadBuilder),
                                                 plot_load_done=printLoads)
                # Choose wish items to put on loads
                for model, crate_type in result:
                    found = False
                    for OnLoad in tempoOnLoad:
                        if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0 and OnLoad.CRATE_TYPE == crate_type:
                            OnLoad.QUANTITY = 0
                            found = True
                            param.AssignedWish.append(OnLoad)
                            OnLoad.EndDate = weekdays(max(0, param.days_to - 1), officialDay=date)
                            SQLPRIORITY.sendToSQL(OnLoad.lineToXlsx(dayTodayComplete))
                            detailed_ws.append(OnLoad.lineToXlsx(dayTodayComplete, filtered=True))
                            DATAWishList.remove(OnLoad)
                            break
                    if not found:
                        GeneralErrors += 'Error in min section: impossible result.\n'
                for wish in tempoOnLoad:  # If it is not on loads, give back inv
                    if wish.QUANTITY > 0:
                        for inv in wish.INV_ITEMS:
                            inv.QUANTITY += 1
                        wish.INV_ITEMS = []

        # Try to Make the maximum number of loads for each P2P

        for param in DATAParams:
            if len(param.LoadBuilder) < param.LOADMAX:  # if we haven't reached the max number of loads
                tempoOnLoad = []
                invData = []

                # Create data table
                for wish in DATAWishList:
                    if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO and wish.QUANTITY > 0:
                        position = 0
                        for Iteration in range(wish.QUANTITY):
                            for It, inv in enumerate(DATAINV[position::]):
                                if EquivalentPlantFrom(inv.POINT, wish.POINT_FROM) and\
                                        inv.MATERIAL_NUMBER == wish.MATERIAL_NUMBER and inv.QUANTITY > 0:
                                    inv.QUANTITY -= 1
                                    wish.INV_ITEMS.append(inv)
                                    position += It
                                    break  # no need to look further
                        if len(wish.INV_ITEMS) < wish.QUANTITY:  # We give back taken inv
                            for invToGiveBack in wish.INV_ITEMS:
                                invToGiveBack.QUANTITY += 1
                            wish.INV_ITEMS = []
                        else:
                            tempoOnLoad.append(wish)
                            invData.append(wish.get_loadbuilder_input_line())

                models_data = loadbuilder_input_dataframe(invData)

                # Create loads
                result = param.LoadBuilder.build(models_data, param.LOADMAX, plot_load_done=printLoads)

                # choose wish items to put on loads
                for model, crate_type in result:
                    found = False
                    for OnLoad in tempoOnLoad:
                        if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0 and OnLoad.CRATE_TYPE == crate_type:
                            OnLoad.QUANTITY = 0
                            found = True
                            param.AssignedWish.append(OnLoad)
                            OnLoad.EndDate = weekdays(max(0, param.days_to - 1), officialDay=date)
                            SQLPRIORITY.sendToSQL(OnLoad.lineToXlsx(dayTodayComplete))
                            detailed_ws.append(OnLoad.lineToXlsx(dayTodayComplete, filtered=True))
                            DATAWishList.remove(OnLoad)
                            break
                    if not found:
                        GeneralErrors += 'Error in max section: impossible result.\n'
                for wish in tempoOnLoad:  # If it is not on loads, give back inv
                    if wish.QUANTITY > 0:
                        for inv in wish.INV_ITEMS:
                            inv.QUANTITY += 1
                        wish.INV_ITEMS = []

        # We set a boolean to false that will indicate if there's anything that has been schedule this date
        consider_date_in_output = False

        # We save the loads created for that day
        for param in DATAParams:

            # print('\nFROM :', param.POINT_FROM, 'TO :', param.POINT_TO)
            # print('DATE OF PLANIFICATION ',  weekdays(-1, officialDay=date))
            # print('DATE OF INVENTORY AVAILABILITY :', date)
            # print('PROJECTED DATE :', weekdays(max(0, param.days_to-1), officialDay=date), '\n')

            SQLLoads.sendToSQL([(param.POINT_FROM, param.POINT_TO, len(param.LoadBuilder),
                                 weekdays(max(0, param.days_to-1), officialDay=date), dayTodayComplete, IsAdhoc)])

            if len(param.LoadBuilder) > 0:
                consider_date_in_output = True
                summary_data.append([param.POINT_FROM, param.POINT_TO, len(param.LoadBuilder),
                                     weekdays(max(0, param.days_to-1), officialDay=date)])

        # Update progress bar and column counter
        progress.update()
        column_counter += int(consider_date_in_output)

    # Loads were made for the next 8 weeks, now we send not assigned wishes to SQL

    for wish in DATAWishList:
        if wish.QUANTITY != wish.ORIGINAL_QUANTITY:
            GeneralErrors += 'Error : this unit was not assigned and has a different quantity; \n {0}'.format(wish.lineToXlsx())
        SQLPRIORITY.sendToSQL(wish.lineToXlsx(dayTodayComplete))

    # We format the summary worksheet and save the workbook and the reference
    fake_columns = [''] * column_counter
    worksheet_formatting(summary_ws, fake_columns, [20] * len(fake_columns))
    reference = [savexlsxFile(wb, saveFolder, dest_filename)]
    wb.close()

    # We create the pivot table and write it in the already saved xlsx file
    workbook = load_workbook(reference[0])
    writer = pd.ExcelWriter(reference[0], engine='openpyxl')
    summary_data = pd.DataFrame(data=summary_data, columns=summary_columns)
    summary_data = pd.pivot_table(summary_data, values='QUANTITY', index=['POINT_FROM', 'SHIPPING_POINT'],
                                  columns=['DEPARTURE_DATE'], aggfunc=np.sum)
    writer.book = workbook
    writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)
    summary_data.to_excel(writer, sheet_name="SUMMARY", header=True)
    writer.save()

    # We send the emails
    send_email(EmailList, dest_filename, 'P2P Forecast is now updated.\n'+GeneralErrors, reference)


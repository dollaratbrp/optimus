"""

Author : Olivier Lefebvre
         Nicolas Raymond

This file contained all main functions related to P2P full process

Last update : 2020-01-14
By : Nicolas Raymond

"""

from LoadBuilder import LoadBuilder, set_trailer_reference
from InputOutput import *
from numpy import savetxt
DATAInclude = []
log_file = None
sharing_points_from = ['4100', '4125']
shared_flatbed_53 = {'QTY': 2, 'POINT_FROM': sharing_points_from}  # Used to keep track of flat53 available
residuals_counter = {}  # Use to keep track of the residuals of min and max among p2p with same POINT TO


class Wish:

    """
    Represents a wish from the wishlist
    """

    def __init__(self, sdn, sin, stn, point_from, shipping_point, div, mat_num, size, length, width, height,
                 stackability, qty, rank, mandatory, overhang, crate_type, valid_from, period_status, is_adhoc=0):

        self.SALES_DOCUMENT_NUMBER = sdn
        self.SALES_ITEM_NUMBER = sin
        self.SOLD_TO_NUMBER = stn
        self.POINT_FROM = point_from
        self.SHIPPING_POINT = shipping_point
        self.DIVISION = div
        self.MATERIAL_NUMBER = mat_num
        self.SIZE_DIMENSIONS = size
        self.LENGTH = length
        self.WIDTH = (1 - (self.SIZE_DIMENSIONS == 'SP2'))*width + (self.SIZE_DIMENSIONS == 'SP2')*self.LENGTH
        self.HEIGHT = height
        self.STACKABILITY = stackability
        self.QUANTITY = qty
        self.RANK = rank
        self.MANDATORY = (mandatory == 'X')
        self.OVERHANG = overhang
        self.IsAdhoc = is_adhoc
        self.CRATE_TYPE = crate_type
        self.VALID_FROM_DATE = valid_from
        self.PERIOD_STATUS = period_status

        # To keep track of inv origins
        self.INV_ITEMS = []
        self.ORIGINAL_QUANTITY = qty

        # If the item is assigned to a Load in excel workbook
        self.Finished = False
        self.EndDate = None

    def lineToXlsx(self, date_today, filtered=False):
        """
        Return a list of all the details needed on a wish to write a line in a .xlsx forecast report
        :param date_today: today's date
        :param filtered: bool to indicate if the list returned will be shortened
        """
        if not filtered:
            return [(self.SALES_DOCUMENT_NUMBER, self.SALES_ITEM_NUMBER, self.SOLD_TO_NUMBER, self.POINT_FROM,
                     self.SHIPPING_POINT, self.DIVISION, self.MATERIAL_NUMBER, self.SIZE_DIMENSIONS,
                     self.LENGTH, self.WIDTH, self.HEIGHT, self.STACKABILITY, self.OVERHANG, self.ORIGINAL_QUANTITY,
                     self.RANK, self.MANDATORY, self.EndDate, date_today, self.IsAdhoc)]
        else:
            return [self.POINT_FROM, self.SHIPPING_POINT, self.DIVISION, self.MATERIAL_NUMBER, self.SIZE_DIMENSIONS,
                    self.ORIGINAL_QUANTITY, self.EndDate]

    def get_loadbuilder_input_line(self):
        """
        Return a list with all details needed on the wish to build the input dataframe of the LoadBuilder
        :return: list
        """
        return [1, self.SIZE_DIMENSIONS, self.LENGTH, self.WIDTH, self.HEIGHT, 1,
                self.CRATE_TYPE, self.STACKABILITY, int(self.MANDATORY), self.OVERHANG]

    def get_log_details(self):
        """
        Returns a list with all strings needed for log file
        :return:
        """
        return ['WISH -> ', ' FROM : ', str(self.POINT_FROM),
                '| TO : ', str(self.SHIPPING_POINT),
                '| MATERIAL_NUMBER : ', str(self.MATERIAL_NUMBER),
                '| SIZE CODE : ', str(self.SIZE_DIMENSIONS),
                '| RANK : ', str(self.RANK), '\n']


class INVObj:

    """
    Represent an inventory line of the inventory data from SQL
    """
    def __init__(self, point, mat_num, qty, date, status):
        self.POINT = point
        self.MATERIAL_NUMBER = mat_num
        self.QUANTITY = qty
        self.DATE = date
        self.STATUS = status
        self.Future = not (weekdays(0) == date)  # if available date is not the same as today
        self.unused = 0  # count the number of skus to display on BOOKED_UNUSED worksheet
        self.SIZE_CODE = None  # Only used to display size_code in BOOKED_UNUSED worksheet
        self.POSSIBLE_PLANT_TO = {}  # Use to show where unbooked quantity could be shipped in BOOKED_UNUSED worksheet

    def lineToXlsx(self, possible_plant_to={}, booked_unused=True):
        """
        Return a list of all the details needed on for BOOKED UNUSED or UNBOOKED worksheet
        """
        if booked_unused:

            # We initialize a list for qty available by possible plant to
            qty_per_plant_to = []

            # We add the quantity in the same order as column names
            for plant in possible_plant_to:
                qty_per_plant_to.append(self.POSSIBLE_PLANT_TO.get(plant, ''))

            return [self.POINT, self.MATERIAL_NUMBER, self.SIZE_CODE, self.unused] + qty_per_plant_to

        else:  # elif unbooked
            return [self.POINT, self.MATERIAL_NUMBER, self.QUANTITY-self.unused]

    def __eq__(self, other):
        """
        Definition of the equality "==" of two INVObj
        :param other: other INVObj
        """
        return (EquivalentPlantFrom(self.POINT, other.POINT) and self.MATERIAL_NUMBER == other.MATERIAL_NUMBER
                and self.Future == other.Future)

    def __iadd__(self, other):
        """
        Definition of "+=" operator for an INVObj
        :param other: other INVObj
        """
        return INVObj(self.POINT, self.MATERIAL_NUMBER, self.QUANTITY+other.QUANTITY, self.DATE, self.STATUS)


class Parameters:

    """
    Represents a line of parameters from the ParameterBox GUI
    """
    def __init__(self, point_from, point_to, loadmin, loadmax, drybox, flatbed, priority, transit,  days_to):

        global shared_flatbed_53
        self.POINT_FROM = point_from
        self.POINT_TO = point_to
        self.LOADMIN = loadmin
        self.LOADMAX = loadmax
        self.ORIGINAL_LOADMAX = loadmax  # Use to reset the in the forecast
        self.TRANSIT = transit
        self.PRIORITY = priority
        self.days_to = days_to
        self.AssignedWish = []
        self.DRYBOX = drybox

        # Trailer quantities initialization
        if self.POINT_FROM in shared_flatbed_53['POINT_FROM']:
            self.FLATBED_48 = max(flatbed - shared_flatbed_53['QTY'], 0)
            self.FLATBED_53 = min(flatbed, shared_flatbed_53['QTY'])
        else:
            self.FLATBED_48 = flatbed
            self.FLATBED_53 = 0

        self.LoadBuilder = self.new_loadbuilder()

    def new_loadbuilder(self):
        """"
        Reset the loadBuilder
        """
        dict_of_qty = {'DRYBOX': self.DRYBOX, 'FLATBED_48': self.FLATBED_48, 'FLATBED_53': self.FLATBED_53}
        categories, quantities = [], []
        for category, qty in dict_of_qty.items():
            if qty > 0:
                categories.append(category)
                quantities.append(qty)

        trailer_data = get_trailers_data(categories, quantities)
        return LoadBuilder(trailer_data)

    def update_load_builder_trailers_data(self):
        """
        Updates the number of flatbed 53 available for our LoadBuilder
        """
        global shared_flatbed_53
        if self.POINT_FROM in shared_flatbed_53['POINT_FROM']:
            df = self.LoadBuilder.trailers_data
            if len(df.loc[df['CATEGORY'] == 'FLATBED_53', 'QTY'].values) != 0:
                df.loc[df['CATEGORY'] == 'FLATBED_53', 'QTY'] = shared_flatbed_53['QTY']

    def update_flatbed_53(self):
        """
        Updates the number of flatbed 53 left
        :return:
        """
        global shared_flatbed_53
        if self.POINT_FROM in shared_flatbed_53['POINT_FROM']:
            df = self.LoadBuilder.trailers_data
            if len(df.loc[df['CATEGORY'] == 'FLATBED_53', 'QTY'].values) != 0:
                shared_flatbed_53['QTY'] = df.loc[df['CATEGORY'] == 'FLATBED_53', 'QTY'].values[0]

    def update_max(self):
        """
        Updates our number of LOADMIN and LOADMAX that we can do according to the residuals from other p2ps
        with the same POINT TO
        """
        global residuals_counter
        global sharing_points_from

        if self.POINT_FROM in sharing_points_from:
            self.LOADMAX += residuals_counter.get(self.POINT_TO, 0)

    def add_residuals(self):
        """
        Add residuals to the residuals counter (used when we try to satisfy max and it's not possible)
        """

        global residuals_counter
        global sharing_points_from

        if self.POINT_FROM in sharing_points_from:

            # We save the number of loads done
            nb_of_loads_done = len(self.LoadBuilder)

            # If we're satisfying the maximum and the number of loads is less than the maximum
            if nb_of_loads_done <= self.LOADMAX:

                # We save the residual number
                residual = self.LOADMAX - nb_of_loads_done

                # We update our LOADMAX attribute
                self.LOADMAX -= residual

                # We send it to our residuals counter
                value_in_place = residuals_counter.setdefault(self.POINT_TO, residual)
                if value_in_place != residual:
                    residuals_counter[self.POINT_TO] = residual

    def remove_residuals(self, qty_to_remove):
        """
        Removes number of new loads done from residuals counter when we are satisfying minimums in the process

        :param qty_to_remove: number of loads to remove from residuals counter
        """
        if self.POINT_FROM in sharing_points_from:

            # We update the load max attribute
            self.LOADMAX += qty_to_remove

            # We reduce the number of residuals by qty_to_remove (or set it to 0)
            residuals_counter[self.POINT_TO] -= qty_to_remove

    def reset(self):
        """
        Reset the number of loadmax and the loadbuilder
        """
        self.LOADMAX = self.ORIGINAL_LOADMAX
        self.LoadBuilder = self.new_loadbuilder()

    def build_loads(self, loadbuilder_input, ranking, temporary_on_load, max_load, print_loads=False, **kwargs):

        """
        Holds all procedures linked to the loadbuilding of a plant to plant

        :param loadbuilder_input: list of lists that will be used to build loadbuilder input dataframe
        :param ranking: dictionary with rankings of crates
        :param temporary_on_load: list of wishes that have been send to create loads
        :param max_load: maximum number of loads to do
        :param print_loads: bool indicating if we should plot the loads done
        """

        # Construction of the data frame which we'll send to the LoadBuilder of our parameters object (p2p)
        input_dataframe = loadbuilder_input_dataframe(loadbuilder_input)

        # We update the trailers data frame of the LoadBuilder associated to the p2p
        self.update_load_builder_trailers_data()

        # We add wishes sent and other info to log file
        add_separator_line('-')
        log_file.writelines(['\n\n\n', 'P2P : {} to {}'.format(str(self.POINT_FROM), str(self.POINT_TO)), '\n'])
        log_file.writelines(['CURRENT MIN : ', str(self.LOADMIN), '| CURRENT MAX : ', str(self.LOADMAX), '\n\n'])
        log_file.writelines(['*** {} wishes available for loads *** '. format(str(len(temporary_on_load))), '\n'])
        for wish in temporary_on_load:
            log_file.writelines(wish.get_log_details())

        log_file.writelines(['\n\n', '*** LOADBUILDER INPUT DATAFRAME *** ', '\n\n'])
        # log_file.writelines([column+' ' for column in input_dataframe.columns])
        # log_file.write('\n')
        for index, row in input_dataframe.iterrows():
            log_file.write('    ')
            log_file.writelines([str(value)+' ' for value in row])
            log_file.write('\n')

        last_number_of_loads = len(self.LoadBuilder)

        # We build loads
        result = self.LoadBuilder.build(input_dataframe, max_load, ranking=ranking, plot_load_done=print_loads)

        # We write the number of loads done
        log_file.writelines(['\n\n', 'NUMBER OF LOADS DONE : {}'.format(str(len(self.LoadBuilder) - last_number_of_loads))])

        # We update the number of common flatbed 53
        self.update_flatbed_53()

        # Choose which wish to send in load based on selected crates and priority order
        link_load_to_wishes(result, temporary_on_load, self, **kwargs)

        # We add a separator line to log file
        add_separator_line('-')

    def get_nb_of_units(self):
        """
        Returns the number of units associated for the p2p
        """
        return self.LoadBuilder.number_of_units()

    def save_full_process_results(self, history_sql_connection, approved_ws_data, sap_input_ws_data, process_date,
                                  saving_path):
        """
        Saves this plant to plant results in three different output format and save pictures of loads done
        1 - Directly in sql history table
        2 - In a list that will be grouped later and used for the APPROVED Worksheet
        3 - In another list that will be used for SAP INPUT worksheet

        :param history_sql_connection: SQLconnection to the history table
        :param approved_ws_data: list that will contain APPROVED worksheet output
        :param sap_input_ws_data: list that will contain SAP INPUT worksheet output
        :param process_date: date when the process was executed
        :param saving_path: path where we'll save the loads pictures
        """
        # We sort trailers with their nb of mandatory and average ranking
        self.LoadBuilder.trailers_done.sort(key=lambda t: (t.nb_of_mandatory(), t.average_ranking()))

        # We init a counter for loads done in this p2p
        load_number = 0

        for trailer in self.LoadBuilder.trailers_done:  # For all loads

            # We update the load number and save picture associated to the current trailer
            load_number += 1
            trailer.plot_load(saving_path=saving_path + '_' + self.POINT_FROM + '_' + self.POINT_TO + '_' +
                              str(load_number))

            for stack in trailer.load:                  # For all stacks in the current load
                for size_dimension in stack.models:     # For all items (size dimension) in the current stack

                    for wish in [w for w in self.AssignedWish if not w.Finished]:

                        if wish.SIZE_DIMENSIONS == size_dimension:

                            # We change the status of the wish
                            wish.Finished = True

                            # We create the line of values to send to sql
                            sql_line = [(self.POINT_FROM, self.POINT_TO, load_number,
                                         wish.MATERIAL_NUMBER, wish.ORIGINAL_QUANTITY,
                                         size_dimension, wish.SALES_DOCUMENT_NUMBER,
                                         wish.SALES_ITEM_NUMBER,
                                         wish.SOLD_TO_NUMBER, process_date)]

                            # We send a line of values to the "APPROVED" worksheet
                            approved_ws_data.append([self.POINT_FROM, self.POINT_TO, load_number,
                                                     trailer.category, round(trailer.length_used/12, 1),
                                                     wish.MATERIAL_NUMBER, wish.ORIGINAL_QUANTITY, size_dimension])

                            # We append a line of data for the "SAP INPUT" worksheet to the list concerned
                            sap_input_ws_data.append([self.POINT_FROM, self.POINT_TO,
                                                      wish.MATERIAL_NUMBER, wish.ORIGINAL_QUANTITY])

                            # We send the sql line to the table concerned
                            history_sql_connection.sendToSQL(sql_line)

                            # We break our research
                            break


class NestedSourcePoints:
    def __init__(self, point_source, point_include):
        self.source = point_source
        self.include = point_include


class FakeLogFile:
    """
    Fake log file that does not execute any task when calling write and write lines on it
    Avoid presence of large number of "if" in the P2P functions
    """
    def __init__(self):
        pass

    def write(self, string):
        pass

    def writelines(self, list_of_strings):
        pass


def clean_p2p_history(expiration_date):
    """
    Delete all rows where import date was set before expiration date

    :param expiration_date: date that determine if a row is expired or not
    :return connection to the table
    """
    # We initialize the connection the the SQL table that will receive the results
    table_header = 'POINT_FROM,SHIPPING_POINT,LOAD_NUMBER,MATERIAL_NUMBER,QUANTITY,SIZE_DIMENSIONS,' \
                   'SALES_DOCUMENT_NUMBER,SALES_ITEM_NUMBER,SOLD_TO_NUMBER,IMPORT_DATE'

    connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_HISTORICAL', headers=table_header)

    connection.deleteFromSQL("IMPORT_DATE < " + "'" + str(expiration_date) + "'")

    return connection


def get_emails_list(project_name):
    """
    Recuperates the list of email associate to the project
    :param project_name: One among 'P2P, 'FORECAST and 'FASTLOADS'
    :return: list of email addresses
    """
    email_connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning',
                                     'OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS', headers='EMAIL_ADDRESS')

    email_query = """ SELECT distinct [EMAIL_ADDRESS]
                 FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS]
                 WHERE PROJECT = """ + "'"+project_name+"'"

    # GET SQL DATA
    email_data = email_connection.GetSQLData(email_query)
    return [email_address for sublist in email_data for email_address in sublist]


def get_p2p_order(sql_connection=None):
    """
    Recuperates the plant to plant in the order in which they must be for the written of the results
    :param sql_connection: already established connection with the sql table that has the data
    :return: list of lists with point_from and point_to
    """

    if sql_connection is None:
        headers = 'POINT_FROM,POINT_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,TRANSIT,PRIORITY_ORDER,SKIP'
        sql_connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS', headers=headers)

    p2p_order_query = """ SELECT distinct  [POINT_FROM],[POINT_TO]
                      FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
                      where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
                      and SKIP = 0
                      order by [POINT_FROM],[POINT_TO]
                    """
    return sql_connection.GetSQLData(p2p_order_query)


def get_missing_p2p():
    """
    Recuperates the p2p that are in the wishlist but are absent in the ParametersBox
    :return: list of lists with point_from and point_to
    """
    connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PRIORITY', headers='')
    query = """SELECT DISTINCT [POINT_FROM]
                      ,[SHIPPING_POINT]
                  FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY] 
                  WHERE CONCAT(POINT_FROM,SHIPPING_POINT) not in (
                    SELECT DISTINCT CONCAT([POINT_FROM],[POINT_TO])
                    FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS] WHERE IMPORT_DATE = (SELECT max(IMPORT_DATE) 
                    FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]) )
                    AND [POINT_FROM] <>[SHIPPING_POINT] 
                """
    return connection.GetSQLData(query)


def shipping_point_names():
    """
    Returns dictionary with names associated to each shipping_points. Useful for outputs
    """

    header = 'SHIPPING_PLANT,SHIPPING_POINT,DESCRIPTION,SOLD_TO_NUMBER'
    shipping_points_connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning',
                                               'BP_CONFIG_SHIPPING_PLANT_TO_SHIPPING_POINT', headers=header)
    query = """SELECT [SHIPPING_POINT], [DESCRIPTION]
                           FROM [Business_Planning].[dbo].[BP_CONFIG_SHIPPING_PLANT_TO_SHIPPING_POINT]
                        """
    shipping_points = {}
    data = shipping_points_connection.GetSQLData(query)
    for lines in data:
        shipping_points[lines[0]] = lines[1]

    return shipping_points


def get_parameter_grid(forecast=False, parameter_box_output=False):
    """
    Recuperates the ParameterBox data from SQL

    :return: list of Parameters and the established SQL connection
    """

    if forecast:
        table = 'OTD_1_P2P_F_FORECAST_PARAMETERS'
    else:
        table = 'OTD_1_P2P_F_PARAMETERS'

    if parameter_box_output:
        headers = "point_FROM,point_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,PRIORITY_ORDER,TRANSIT,DAYS_TO,SKIP,IMPORT_DATE"
        order = """[POINT_TO] ,[POINT_FROM]"""
        skip_column = ",[SKIP]"
        skip_filter = ""
    else:
        headers = 'POINT_FROM,POINT_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,TRANSIT,PRIORITY_ORDER,SKIP'
        order = "PRIORITY_ORDER"
        skip_column = ""
        skip_filter = "and SKIP = 0"

    connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', table, headers=headers)

    query = """ SELECT [POINT_FROM]
                      ,[POINT_TO]
                      ,[LOAD_MIN]
                      ,[LOAD_MAX]
                      ,[DRYBOX]
                      ,[FLATBED]
                      ,[PRIORITY_ORDER]
                      ,[TRANSIT] """ + skip_column + """
                      ,DAYS_TO
                  FROM """ + table + """ 
                  where IMPORT_DATE = (select max(IMPORT_DATE) from """ + table + """ ) """ + skip_filter + """
                  order by """ + order
    print(query)
    # GET SQL DATA
    data = connection.GetSQLData(query)
    if parameter_box_output:
        return [line for line in data], connection
    else:
        return [Parameters(*line) for line in data], connection


def reset_flatbed_53():
    """
    Resets the number of flatbed 53 available
    """
    global shared_flatbed_53
    shared_flatbed_53 = {'QTY': 2, 'POINT_FROM': ['4100', '4125']}


def reset_residuals_counter():
    """
    Resets the residuals counter
    """
    global residuals_counter
    residuals_counter = {}


def get_wish_list(forecast=False):

    """
    Recuperates the whish list from SQL
    :param : bool indicating if we want the wishlist for the forecast
    :return : list of Wish object
    """

    wishlist_headers = 'SALES_DOCUMENT_NUMBER,SALES_ITEM_NUMBER,SOLD_TO_NUMBER,POINT_FROM,' \
                       'SHIPPING_POINT,DIVISION,MATERIAL_NUMBER,Size_Dimension,Lenght,Width,' \
                       'Height,stackability,Quantity,Priority_Rank,X_IF_MANDATORY, METAL_WOOD'

    wishlist_connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning',
                                        'OTD_2_PRIORITY_F_P2P', headers=wishlist_headers)
    if forecast:
        parameters_table = '[Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS]'
        period_status = "[PERIOD_STATUS] in ('" + 'P2P' + "','"+'FCST' + "')"  # MUST BE CHANGED !!!!
    else:
        parameters_table = '[Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]'
        period_status = '[PERIOD_STATUS] = ' + "'" + 'P2P' + "'"

    query = """SELECT  [SALES_DOCUMENT_NUMBER]
                      ,[SALES_ITEM_NUMBER]
                      ,[SOLD_TO_NUMBER]
                      ,[POINT_FROM]
                      ,[SHIPPING_POINT]
                      ,[DIVISION]
                      ,RTRIM([MATERIAL_NUMBER])
                      ,RTRIM([Size_Dimension])
                      ,convert(int,CEILING([Length]))
                      ,convert(int,CEILING([Width]))
                      ,convert(int,CEILING([Height]))
                      ,convert(int,[stackability])
                      ,[Quantity]
                      ,[Priority_Rank]
                      ,[X_IF_MANDATORY]
                      ,[OVERHANG]
                      ,[METAL_WOOD]
                      ,[valid_from_date]
                      ,[PERIOD_STATUS]
                  FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY]
                  WHERE [POINT_FROM] <> [SHIPPING_POINT] 
                  AND Length <> 0 and Width <> 0 AND Height <> 0 and """ + period_status + """
                  AND concat(POINT_FROM,SHIPPING_POINT) in 
                  (select distinct concat([POINT_FROM],[POINT_TO]) from """ + parameters_table + """
                  where IMPORT_DATE = (select max(IMPORT_DATE) from """ + parameters_table + """) and SKIP = 0)
                  order by Priority_Rank 
                """

    data = wishlist_connection.GetSQLData(query)
    return [Wish(*line) for line in data]


def get_inventory_and_qa():

    """Recuperates the inventory and QA HOLD data from SQL"""

    # We first get the inventory
    connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning',
                               'OTD_1_P2P_F_INVENTORY', headers='')

    inventory_query = """ SELECT *
        FROM (
        SELECT DISTINCT SHIPPING_POINT
          ,RTRIM([MATERIAL_NUMBER]) as MATERIAL_NUMBER
          ,sum(tempo.QUANTITY) as [QUANTITY]
          ,CONVERT(DATE,GETDATE()) as [AVAILABLE_DATE]
          ,'INVENTORY' as [STATUS]
          FROM(
      SELECT  [SHIPPING_POINT]
          ,[MATERIAL_NUMBER]
          ,[QUANTITY]
          ,[AVAILABLE_DATE]
          ,[STATUS]
      FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
      WHERE STATUS = 'INVENTORY'
      UNION 
      (SELECT [SHIPPING_POINT]
          ,[MATERIAL_NUMBER]
          ,[QUANTITY]
          ,GETDATE() as [AVAILABLE_DATE]
          ,'INVENTORY' as [STATUS]
           FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
      WHERE STATUS in ('QA HOLD') and AVAILABLE_DATE between convert(DATE,GETDATE()-1) and GETDATE())) as tempo
     GROUP BY SHIPPING_POINT 
          ,[MATERIAL_NUMBER]
          ,  [AVAILABLE_DATE]
          , [STATUS] ) as tempo2
     
     ORDER BY SHIPPING_POINT, MATERIAL_NUMBER
                        """

    data = connection.GetSQLData(inventory_query)
    inventory = [INVObj(*line) for line in connection.GetSQLData(inventory_query)]

    # We then take the QA HOLD
    qa_query = """ SELECT  [SHIPPING_POINT]
                      ,RTRIM([MATERIAL_NUMBER]) as MATERIAL_NUMBER
                      ,convert(int,[QUANTITY]) as QUANTITY
                      ,convert (DATE,[AVAILABLE_DATE]) as AVAILABLE_DATE
                      ,[STATUS]
                  FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
                  where STATUS = 'QA HOLD'
                  and AVAILABLE_DATE = (case when DATEPART(WEEKDAY,getdate()) = 6 then convert(DATE,GETDATE()+3) else convert(DATE,GETDATE() +1) end)
                    """

    data = connection.GetSQLData(qa_query)

    # we want the QA at the end of inv list, so the skus in QA will be the last to be chose
    for line in data:
        inventory.append(INVObj(*line))  # add QA HOLD with inv

    official_inventory = adjust_inventory(inventory)

    return official_inventory


def adjust_inventory(original_inventory):
    """
    Group common inventory together
    :param original_inventory: list of INVObj
    """

    # We initialize a list of official INVObj that we'll keep
    official_inventory = []

    # We group common INVobj together
    while len(original_inventory) > 0:

        # We save the current object we're looking at
        current_obj = original_inventory[0]

        # original_qty = current_obj.QUANTITY

        # We initialize an empty list that will contain index of INVObj
        indexes = []

        # We go through all the inventory to sum qty of equivalent object
        for i in range(1, len(original_inventory)):

            if current_obj == original_inventory[i]:
                current_obj += original_inventory[i]

                indexes.append(i)

        # We remove INVObjs which the inventory was included in he current INVObj
        remove_indexes_from_list(original_inventory, indexes)

        # If the qty of the current INVObj is greater than 0 we add it to the official inventory
        if current_obj.QUANTITY > 0:
            official_inventory.append(current_obj)

        original_inventory.pop(0)

    return official_inventory


def get_nested_source_points(l):

    """
    Gets all information on which plant inventory is included in which other (from SQL)
    and push it in l

    :param l: list that will contained the NestedSourcePoints
    """

    connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning',
                               'OTD_1_P2P_D_INCLUDED_INVENTORY', headers='')

    query = """select SHIPPING_POINT_SOURCE ,SHIPPING_POINT_INCLUDE
               from OTD_1_P2P_D_INCLUDED_INVENTORY
             """
    data = connection.GetSQLData(query)

    for line in data:
        l.append(NestedSourcePoints(*line))


def get_trailers_data(category_list=[], qty_list=[]):
    """
    Gets the trailers data from SQL

    :param category_list: list with the categories of trailer which we need the data
    :param qty_list: list with quantities of each category of trailer

    *** IMPORTANT *** : Categories and quantities must be in the same order
    :return: list of lists for every line of data
    """
    # Initialization of column names for data that will be sent to LoadBuilder
    columns = ['QTY', 'CATEGORY', 'LENGTH', 'WIDTH', 'HEIGHT', 'OVERHANG']

    # Connection to SQL database that contains data needed
    sql_connect = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_TRUCK_PARAMETERS')

    # We look category_list size
    if len(category_list) != len(qty_list):
        raise Exception("\nLengths of the two inputs are not the same")

    elif len(category_list) == 0:
        end_of_query = ""

    elif len(category_list) == 1:
        end_of_query = """WHERE RTRIM([TYPE]) = """ + "'" + str(category_list[0]) + "'"

    else:
        end_of_query = """WHERE RTRIM([TYPE]) in """ + str(tuple(category_list))

    # Writing of our query
    sql_query = """SELECT RTRIM([TYPE])
    ,[LENGTH]
    ,[INTERIOR_WIDTH]
    ,[HEIGHT]
    ,[OVERHANG]
    FROM [Business_Planning].[dbo].[OTD_1_P2P_F_TRUCK_PARAMETERS] """ + end_of_query

    # Retrieve the data
    data = sql_connect.GetSQLData(sql_query)

    # Add each qty at the beginning of the good line of data
    if len(category_list) != 0:
        for line in data:
            index_of_qty = category_list.index(line[0])
            line.insert(0, qty_list[index_of_qty])
    else:
        for line in data:
            line.insert(0, 0)

    return pd.DataFrame(data=data, columns=columns)


def find_perfect_match(Wishes, Inventory, Parameters):

    """
    Finds perfect match between wishes of the wish list, inventory available and p2p available in the parameter box
    :param Wishes: List of wishes (list of WishlistObj)
    :param Inventory: List of INVobj
    :param Parameters: List of Parameters
    :return: List of whishes approved
    """
    # We indicate in the log file in which part of the process we are
    add_separator_line()
    log_file.writelines(['\n\n\n', 'PERFECT MATCH', '\n\n\n'])

    # We initialize a list that will contain all wish approved
    ApprovedWish = []

    # We iterate through wishlist to keep the priority order
    for wish in Wishes:

        position = 0  # To not loop through all inv Data at each iteration

        # For all unit needed to fulfill our wish
        for unit_needed in range(wish.QUANTITY):

            # For all pairs of (index, INVobj) of our list of INVobj
            for It, inv in enumerate(Inventory[position::]):

                # If the wish and inventory POINT FROM are equivalent and material number are the same and the quantity
                # attribute of the INVobj is greater than 0 (there's still units of this SKU in inventory)
                if EquivalentPlantFrom(inv.POINT, wish.POINT_FROM) and wish.MATERIAL_NUMBER == inv.MATERIAL_NUMBER \
                        and inv.QUANTITY > 0:

                    # If the inventory object (INVobj) will be available later than today
                    if inv.Future:  # QA of tomorrow, need to look if load is for today or later

                        inventory_to_take = False

                        # We loop through our parameters obj through DATAParams list
                        # (We can see it as looping through lines of parameter box)
                        for p2p in Parameters:

                            # If p2p point_from and point_to are corresponding with the wish
                            if wish.POINT_FROM == p2p.POINT_FROM and wish.SHIPPING_POINT == p2p.POINT_TO \
                                    and p2p.days_to > 0:

                                # We must take the item
                                inventory_to_take = True
                                break

                        # If we decided to take the item from the inventory
                        if inventory_to_take:

                            # We decrease its quantity and add the INVobj to the list of inventory items of the wish
                            inv.QUANTITY -= 1
                            wish.INV_ITEMS.append(inv)
                            position += It
                            break  # no need to look further

                    # Else, the INVobj is available today
                    else:

                        # we give the inv to the wish item
                        inv.QUANTITY -= 1
                        wish.INV_ITEMS.append(inv)
                        position += It
                        break  # no need to look further

        # We give back taken inv if there is not enough units to fulfill a wish (build a crate)
        if len(wish.INV_ITEMS) < wish.QUANTITY:
            for invToGiveBack in wish.INV_ITEMS:
                invToGiveBack.QUANTITY += 1
            wish.INV_ITEMS = []
        else:
            ApprovedWish.append(wish)

    # We write approved wishes in log file
    log_file.writelines(['*** {} PERFECT MATCH FOUND *** '.format(str(len(ApprovedWish))), '\n'])
    for wish in ApprovedWish:
        log_file.writelines(wish.get_log_details())

    return ApprovedWish


def perfect_match_loads_construction(Parameters, ApprovedWishes, print_loads=False, **kwargs):
    """
    Builds loads after perfect match

    :param Parameters: List with Parameters objects
    :param ApprovedWishes: List of wishes approved
    :param print_loads: bool indicating if we must plot loads done
    """

    # We write the current step we're doing in the log file
    add_separator_line()
    log_file.writelines(['\n\n\n', 'PERFECT MATCH LOADS CONSTRUCTION', '\n\n\n'])

    for param in Parameters:  # for all P2P in parameters

        # We update LOADMIN and LOADMAX attribute
        param.update_max()

        # Initialization of empty list
        temporary_on_load = []  # List to remember the INVobjs that will be sent to the LoadBuilder
        loadbuilder_input = []  # List that will contain the data to build the frame we'll send to the LoadBuilder

        # Initialization of an empty ranking dictionary
        ranking = {}

        # We loop through our wishes list
        for wish in ApprovedWishes:

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

        param.build_loads(loadbuilder_input, ranking, temporary_on_load, param.LOADMAX, print_loads=print_loads, **kwargs)
        param.add_residuals()

    # Store unallocated units in inv pool
    throw_back_to_pool(ApprovedWishes)


def satisfy_max_or_min(Wishes, Inventory, Parameters, satisfy_min=True, print_loads=False, **kwargs):
    """
    Attributes wishes wisely among p2p'S in Parameters list in order to satisfy their min or their max value

    :param Wishes: List of wishes (list of WishlistObj)
    :param Inventory: List of INVobj
    :param Parameters: List of Parameters
    :param satisfy_min: (bool) if false -> we want to satisfy the max
    :param print_loads: (bool) indicates if we plot each load or not
    """
    global residuals_counter

    # We look if we're distributing leftovers or processing to normal stack packing
    leftover_distribution = kwargs.get('leftovers', False)

    # We write the current step we're doing in the log file
    add_separator_line()
    if leftover_distribution:
        log_file.writelines(['\n\n\n', 'LEFTOVER DISTRIBUTION', '\n\n\n'])
    elif satisfy_min:
        log_file.writelines(['\n\n\n', 'SATISFY MINIMUMS', '\n\n\n'])
    else:
        log_file.writelines(['\n\n\n', 'SATISFY MAXIMUMS', '\n\n\n'])

    # We save a "trigger" integer value indicating if we want to satisfy min or max
    check_min = int(satisfy_min)  # Will be 1 if we want to satisfy min and 0 instead

    # We update LoadBuilder class attribute plc_lb depending on the situation
    LoadBuilder.plc_lb = 0.74 * check_min + (1 - check_min) * 0.80

    # For each parameters in Parameters list
    for param in Parameters:

        # We update LOADMAX attribute
        if not satisfy_min:
            param.update_max()

        # We save the current number of loads done
        nb_loads_done = len(param.LoadBuilder)

        if nb_loads_done < (check_min*param.LOADMIN + (1-check_min)*param.LOADMAX) or leftover_distribution:

            # Initialization of empty list
            temporary_on_load = []  # List to remember the INVobj that will be sent to the LoadBuilder
            load_builder_input = []  # List that will contain the data to build the frame we'll send to the LoadBuilder

            # Initialization of an empty ranking dictionary
            ranking = {}

            # We loop through our wishes list
            for wish in Wishes:

                # If the wish is not fulfilled and his POINT FROM and POINT TO are corresponding with the param (p2p)
                if wish.QUANTITY > 0 and wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO:

                    position = 0

                    # We look if there's inventory available to satisfy each unit needed for our wish
                    for unit_needed in range(wish.QUANTITY):

                        # For all pairs of (index, INVObj) of our list of INVObj
                        for It, inv in enumerate(Inventory[position::]):
                            if EquivalentPlantFrom(inv.POINT, wish.POINT_FROM) and\
                                    inv.MATERIAL_NUMBER == wish.MATERIAL_NUMBER and inv.QUANTITY > 0 and\
                                    (not inv.Future or (inv.Future and param.days_to > 0)):

                                inv.QUANTITY -= 1
                                wish.INV_ITEMS.append(inv)
                                position += It
                                break  # no need to look further

                    # We give back taken inv if there is not enough units to fulfill a wish (build a crate)
                    if len(wish.INV_ITEMS) < wish.QUANTITY:  # We give back taken inv
                        for invToGiveBack in wish.INV_ITEMS:
                            invToGiveBack.QUANTITY += 1
                        wish.INV_ITEMS = []

                    # If the wish can be satisfied
                    else:
                        temporary_on_load.append(wish)

                        # Here we set QTY and NBR_PER_CRATE to 1 because each line of the wishlist correspond to
                        # one crate and not one unit! Must be done this way to avoid having getting to many size_code
                        # in the returning list of the LoadBuilder
                        load_builder_input.append(wish.get_loadbuilder_input_line())

                        # We add the ranking of the wish in the ranking dictionary
                        if wish.SIZE_DIMENSIONS in ranking:
                            ranking[wish.SIZE_DIMENSIONS] += [wish.RANK]
                        else:
                            ranking[wish.SIZE_DIMENSIONS] = [wish.RANK]

            # We save the maximum number of loads that can be done
            if leftover_distribution:
                max_load = 0
            else:
                max_load = (check_min * min(param.LOADMIN, residuals_counter.get(param.POINT_TO, param.LOADMIN)) + (1 - check_min) * param.LOADMAX)

            # We build loads:
            param.build_loads(load_builder_input, ranking, temporary_on_load, max_load, print_loads=print_loads, **kwargs)

            # Store unallocated units in inv pool
            throw_back_to_pool(temporary_on_load)

        # If we are satisfying max
        if not satisfy_min:
            param.add_residuals()

        # Else we're satisfying min
        else:
            # If there's new load done we remove them from residuals counter
            nb_new_loads = (len(param.LoadBuilder) - nb_loads_done)
            if nb_new_loads > 0:
                param.remove_residuals(nb_new_loads)


def distribute_leftovers(Wishes, Inventory, Parameters):

    """
    Distributes leftover crates among packed trailers that aren't completely filled

    :param Wishes: List of wishes (list of WishlistObj)
    :param Inventory: List of INVobj
    :param Parameters: List of Parameters
    :return:
    """
    # We set all LoadBuilders attribute "patching_activated" to True
    LoadBuilder.patching_activated = True

    # We distribute leftover crates among trailers that aren't completely filled
    satisfy_max_or_min(Wishes, Inventory, Parameters, satisfy_min=False, leftovers=True)


def loadbuilder_input_dataframe(data):
    """
    Builds the appropriate pandas data frame needed as input of the LoadBuilder
    :param data: List of lists containing every line of data for our frame
    :return: pandas dataframe
    """

    # Creation of the data frame
    input_frame = pd.DataFrame(data=data, columns=['QTY', 'MODEL', 'LENGTH', 'WIDTH',
                                                   'HEIGHT', 'NBR_PER_CRATE', 'CRATE_TYPE',
                                                   'STACK_LIMIT', 'NB_OF_X', 'OVERHANG'])

    # Group by to sum quantity
    input_frame = input_frame.groupby(['MODEL', 'LENGTH', 'WIDTH', 'HEIGHT',
                                       'NBR_PER_CRATE', 'CRATE_TYPE', 'STACK_LIMIT', 'OVERHANG']).sum()

    # Reformatting of the new object as a standard data frame
    input_frame = input_frame.reset_index()

    return input_frame


def link_load_to_wishes(loadbuilder_output, available_wishes, p2p, **kwargs):
    """
    Choose which wishes to link with the load based on selected crates and priority order
    :param loadbuilder_output: LoadBuilder output (list of tuples with size_code and crate type)
    :param available_wishes: List of wishes that were temporary assigned to the load
    :param p2p : plant to plant for which we built the load (object of class Parameters)
    """
    # We look if there's any function to save wish assignment and a date that are also passed as parameter
    save_wish_assignment = kwargs.get('assignment_function', None)
    inventory_availability_date = kwargs.get('inventory_available_date', None)

    # If we got both parameters we looked for, it means we're linking load to wishes in the forecast process
    forecast_process = save_wish_assignment is not None and inventory_availability_date is not None

    log_file.writelines(['\n\n\n', '*** {} wishes assigned *** '.format(str(len(loadbuilder_output))), '\n'])
    for model, crate_type in loadbuilder_output:
        found = False
        for wish in available_wishes:
            if wish.SIZE_DIMENSIONS == model and wish.QUANTITY > 0 and crate_type == wish.CRATE_TYPE:
                log_file.writelines(wish.get_log_details())
                wish.QUANTITY = 0
                found = True
                p2p.AssignedWish.append(wish)
                if forecast_process:
                    save_wish_assignment(wish, p2p, inventory_availability_date)
                break
        if not found:
            print('Error in Perfect Match: impossible result.\n')


def throw_back_to_pool(wishes):
    """
    Get inventory back to inventory pool if it wasn't used in the current load
    :param wishes: list of wish that weren't fulfilled
    """
    for wish in wishes:
        if wish.QUANTITY > 0:
            for inv in wish.INV_ITEMS:
                inv.QUANTITY += 1
            wish.INV_ITEMS = []


def compute_booked_unused(wishlist, inventory, parameters):
    """
    Match unused inventory to items remaining in the wish list
    :param wishlist: List of wishes
    :param inventory: List of INVObj
    :param parameters: list of Parameters (plant to plant)
    :return: List of distinct plant to where remaining items could be shipped
    """
    # We initialize an empty set that will contain all possible plant to where booked unused could be shipped
    possible_plant_to = set()

    # Assign left inventory to wishList, to look for booked but unused
    for wish in wishlist:

        # We look for any conflict
        if wish.QUANTITY == 0 and not wish.Finished:
            print('Error with wish: ', wish.lineToXlsx())

        # If the wish was not fulfilled
        if wish.QUANTITY > 0:

            # We set a position index to avoid going through all the inventory every time
            position = 0

            # For every unit need to fulfill the wish
            for i in range(wish.QUANTITY):

                # For index and INVobj in the inventory from position index
                for j, inv in enumerate(inventory[position::]):
                    if EquivalentPlantFrom(inv.POINT, wish.POINT_FROM) and \
                            wish.MATERIAL_NUMBER == inv.MATERIAL_NUMBER and inv.QUANTITY - inv.unused > 0:

                        # QA of tomorrow, need to look if load is for today or later
                        if inv.Future:
                            inventory_to_take = False
                            for param in parameters:
                                if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO \
                                        and param.days_to > 0:
                                    inventory_to_take = True
                                    break
                            if inventory_to_take:
                                plant_to = wish.SHIPPING_POINT
                                possible_plant_to.add(plant_to)
                                inv.unused += 1
                                inv.POSSIBLE_PLANT_TO[plant_to] = inv.POSSIBLE_PLANT_TO.get(plant_to, 0) + 1
                                inv.SIZE_CODE = wish.SIZE_DIMENSIONS
                                position += j
                                break  # no need to look further
                        else:
                            plant_to = wish.SHIPPING_POINT
                            possible_plant_to.add(plant_to)
                            inv.unused += 1
                            inv.POSSIBLE_PLANT_TO[plant_to] = inv.POSSIBLE_PLANT_TO.get(plant_to, 0) + 1
                            inv.SIZE_CODE = wish.SIZE_DIMENSIONS
                            position += j
                            break  # no need to look further

    return list(possible_plant_to)


def build_forecast_output_sending_function(wishlist, priority_table_connection,
                                           forecast_running_date, detailed_worksheet):
    """

    Builds the function need to throw output of forecast correctly is function "link_load_to_wishes"

    :param wishlist: complete list of wishes
    :param priority_table_connection: SQLconnection object connect to OTD_1_P2P_F_FORECAST_PRIORITY table
    :param forecast_running_date: date at which the forecast process is running
    :param detailed_worksheet: worksheet on which we write the detailed results of forecast
    :return: A function
    """
    def save_wish_assignment(wish, p2p, inventory_availability_date):
        """
        Saves wish assignement to SQL forecast output table and DETAILED forecast output worksheet

        :param wish: wish that we save the assignment
        :param p2p: plant to plant to which the wish is assigned
        :param inventory_availability_date: date where the inventory to fulfill the wish is available
        """
        wish.EndDate = weekdays(max(0, p2p.days_to - 1), officialDay=inventory_availability_date)
        priority_table_connection.sendToSQL(wish.lineToXlsx(forecast_running_date))
        detailed_worksheet.append(wish.lineToXlsx(forecast_running_date, filtered=True))
        wishlist.remove(wish)

    return save_wish_assignment


def EquivalentPlantFrom(Point1, Point2):
    """" Point1 is shipping_point_from for inv, Point2 is shipping_point_from for wishlist
        Point1 is included in Point2                                                      """
    if Point1 == Point2:
        return True
    else:
        for equiv in DATAInclude:
            if equiv.source == Point2 and equiv.include == Point1:
                return True
    return False


def remove_indexes_from_list(list_to_modify, indexes_list):
    """
    Remove all indexes mentionned from the list

    :param list_to_modify: list from which we'll remove item at the indexes mentionned
    :param indexes_list: list of indexes (list of int)
    """
    indexes_list.sort(reverse=True)
    for i in indexes_list:
        list_to_modify.pop(i)


def create_log_file(path, log_save=True):
    """
    Creates log file to store activities done during the processes
    :param path: path where the file will be stored
    :param log_save: bool indicating if we create a fake log file that doesn't save anything
    """
    global log_file
    if log_save:
        log_file = open(path+'log.txt', 'a')
    else:
        log_file = FakeLogFile()


def write_departure_date(departure_date):
    """
    Writes the departure date in the log file (used in forecast)
    :param departure_date: date to write
    :return:
    """
    global log_file

    # We write the departure date in the log file
    log_file.writelines(['\n\n\n', 'DEPARTURE DATE : ', str(departure_date), '\n\n\n'])


def add_separator_line(symbol='#'):
    """
    Add a line of '#' symbols in the log file
    """
    global log_file
    log_file.writelines(['\n\n'] + [symbol]*80)



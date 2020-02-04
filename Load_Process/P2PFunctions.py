"""

Author : Olivier Lefebvre
         Nicolas Raymond

This file contained all main functions related to P2P full process

Last update : 2020-01-14
By : Nicolas Raymond

"""

from LoadBuilder import LoadBuilder, set_trailer_reference
from InputOutput import *
DATAInclude = []
shared_flatbed_53 = {'QTY': 2, 'POINT_FROM': ['4100', '4125']}


class Wish:

    """
    Represents a wish from the wishlist
    """

    def __init__(self, sdn, sin, stn, point_from, shipping_point, div, mat_num, size, length, width, height,
                 stackability, qty, rank, mandatory, overhang, crate_type, is_adhoc=0):

        self.SALES_DOCUMENT_NUMBER = sdn
        self.SALES_ITEM_NUMBER = sin
        self.SOLD_TO_NUMBER = stn
        self.POINT_FROM = point_from
        self.SHIPPING_POINT = shipping_point
        self.DIVISION = div
        self.MATERIAL_NUMBER = mat_num
        self.SIZE_DIMENSIONS = size
        self.LENGTH = length
        self.WIDTH = width
        self.HEIGHT = height
        self.STACKABILITY = stackability
        self.QUANTITY = qty
        self.RANK = rank
        self.MANDATORY = (mandatory == 'X')
        self.OVERHANG = overhang
        self.IsAdhoc = is_adhoc
        self.CRATE_TYPE = crate_type

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

    def lineToXlsx(self):
        """
        Return a list of all the details needed on a inventory object to write a line in a .xlsx forecast report
        """
        return [self.POINT, self.MATERIAL_NUMBER, self.QUANTITY, self.DATE, self.STATUS]

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
    def __init__(self, point_from, point_to, loadmin, loadmax, drybox, flatbed, transit, priority, days_to):

        global shared_flatbed_53
        self.POINT_FROM = point_from
        self.POINT_TO = point_to
        self.LOADMIN = loadmin
        self.LOADMAX = loadmax
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

        self.LoadBuilder = self.new_LoadBuilder()

    def new_LoadBuilder(self):
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
        Updates the number of flatbet 53 left
        :return:
        """

        global shared_flatbed_53
        if self.POINT_FROM in shared_flatbed_53['POINT_FROM']:
            df = self.LoadBuilder.trailers_data
            if len(df.loc[df['CATEGORY'] == 'FLATBED_53', 'QTY'].values) != 0:
                shared_flatbed_53['QTY'] = df.loc[df['CATEGORY'] == 'FLATBED_53', 'QTY'].values[0]


class NestedSourcePoints:
    def __init__(self, point_source, point_include):
        self.source = point_source
        self.include = point_include


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


def get_parameter_grid():
    """
    Recuperates the ParameterBox data from SQL

    :return: list of Parameters and the established SQL connection
    """
    headers = 'POINT_FROM,POINT_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,TRANSIT,PRIORITY_ORDER,SKIP'

    connection = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS', headers=headers)

    query = """ SELECT  [POINT_FROM]
                      ,[POINT_TO]
                      ,[LOAD_MIN]
                      ,[LOAD_MAX]
                      ,[DRYBOX]
                      ,[FLATBED]
                      ,[TRANSIT]
                      ,[PRIORITY_ORDER]
                      ,DAYS_TO
                  FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
                  where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
                  and SKIP = 0
                  order by PRIORITY_ORDER
                """
    # GET SQL DATA
    data = connection.GetSQLData(query)
    return [Parameters(*line) for line in data], connection


def reset_flatbed_53():
    """
    Resets the number of flatbed 53 available
    """
    global shared_flatbed_53
    shared_flatbed_53 = {'QTY': 2, 'POINT_FROM': ['4100', '4125']}


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
    else:
        parameters_table = '[Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]'

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
                  FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY]
                  WHERE [POINT_FROM] <> [SHIPPING_POINT] AND Length <> 0 and Width <> 0 AND Height <> 0
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

        # We initialize an empty list that will contain index of INVObj
        indexes = []

        # We go through all the inventory to sum qty of equivalent object
        for i in range(1, len(original_inventory)):

            if current_obj == original_inventory[i]:
                current_obj += original_inventory[i]

                indexes.append(i)

        # We remove INVObjs which the inventory was included in he current INVObj
        indexes.sort(reverse=True)
        for i in indexes:
            original_inventory.pop(i)

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

                        InvToTake = False

                        # We loop through our parameters obj through DATAParams list
                        # (We can see it as looping through lines of parameter box)
                        for p2p in Parameters:

                            # If p2p point_from and point_to are corresponding with the wish
                            if wish.POINT_FROM == p2p.POINT_FROM and wish.SHIPPING_POINT == p2p.POINT_TO \
                                    and p2p.days_to > 0:

                                # We must take the item
                                InvToTake = True
                                break

                        # If we decided to take the item from the inventory
                        if InvToTake:

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

    return ApprovedWish


def satisfy_max_or_min(Wishes, Inventory, Parameters, satisfy_min=True, print_loads=False, **kwargs):
    """
    Attributes wishes wisely among p2p'S in Parameters list in order to satisfy their min or their max value

    :param Wishes: List of wishes (list of WishlistObj)
    :param Inventory: List of INVobj
    :param Parameters: List of Parameters
    :param satisfy_min: (bool) if false -> we want to satisfy the max
    :param print_loads: (bool) indicates if we plot each load or not
    """
    # We look if we're distributing leftovers or processing to normal stack packing
    leftover_distribution = kwargs.get('leftovers', False)

    # We save a "trigger" integer value indicating if we want to satisfy min or max
    check_min = int(satisfy_min)  # Will be 1 if we want to satisfy min and 0 instead

    # For each parameters in Parameters list
    for param in Parameters:

        if len(param.LoadBuilder) < (check_min*param.LOADMIN + (1-check_min)*param.LOADMAX) or leftover_distribution:

            # We update LoadBuilder plc_lb depending on the situation
            param.LoadBuilder.plc_lb = 0.75*check_min + (1-check_min)*0.80

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

                        # For all pairs of (index, INVobj) of our list of INVobj
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

            # Construction of the data frame which we'll send to the LoadBuilder of our parameters object (p2p)
            input_dataframe = loadbuilder_input_dataframe(load_builder_input)

            # We update the trailers dataframe of the LoadBuild associated to the p2p
            param.update_load_builder_trailers_data()

            # We save the maximum number of loads that can be done
            if leftover_distribution:
                max_load = 0
            else:
                max_load = (check_min * param.LOADMIN + (1 - check_min) * param.LOADMAX)

            # Construction of loadings
            result = param.LoadBuilder.build(models_data=input_dataframe,
                                             max_load=max_load,
                                             plot_load_done=print_loads,
                                             ranking=ranking)

            # We update the number of common flatbed 53
            param.update_flatbed_53()

            # Choice the wish items to put on loads
            link_load_to_wishes(result, temporary_on_load, param)

            # Store unallocated units in inv pool
            throw_back_to_pool(temporary_on_load)


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


def link_load_to_wishes(loadbuilder_output, available_wishes, p2p):
    """
    Choose which wishes to link with the load based on selected crates and priority order
    :param loadbuilder_output: LoadBuilder output (list of tuples with size_code and crate type)
    :param available_wishes: List of wishes that were temporary assigned to the load
    :param p2p : plant to plant for which we built the load (object of class Parameters)
    """
    for model, crate_type in loadbuilder_output:
        found = False
        for wish in available_wishes:
            if wish.SIZE_DIMENSIONS == model and wish.QUANTITY > 0 and crate_type == wish.CRATE_TYPE:
                wish.QUANTITY = 0
                found = True
                p2p.AssignedWish.append(wish)
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


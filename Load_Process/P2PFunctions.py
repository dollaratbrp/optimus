"""

Author : Olivier Lefebre,
         Nicolas Raymond

This file contained informations on WishlistObj class

Last update : 2020-01-14
By : Nicolas Raymond

"""

from LoadBuilder import LoadBuilder
from InputOutput import *
DATAInclude = []


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


class Parameters:

    """
    Represents a line of parameters from the ParameterBox GUI
    """
    def __init__(self, point_from, point_to, loadmin, loadmax, drybox, flatbed, transit, priority, days_to):
        self.POINT_FROM = point_from
        self.POINT_TO = point_to
        self.LOADMIN = loadmin
        self.LOADMAX = loadmax
        self.DRYBOX = drybox
        self.FLATBED = flatbed
        self.TRANSIT = transit
        self.PRIORITY = priority
        self.days_to = days_to

        self.LoadBuilder = []
        self.AssignedWish = []
        self.new_LoadBuilder()

    def new_LoadBuilder(self):
        """"
        We initialize the loadBuilder
        """
        trailer_data = get_trailers_data(['DRYBOX', 'FLATBED'], [self.DRYBOX, self.FLATBED])
        self.LoadBuilder = LoadBuilder(trailer_data)


class NestedSourcePoints:
    def __init__(self, point_source, point_include):
        self.source = point_source
        self.include = point_include


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
          ,CASE WHEN sum(tempo.[QUANTITY]) <0 then 0 else convert(int,sum(tempo.QUANTITY)) end as [QUANTITY]
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
     WHERE QUANTITY > 0
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
                  where STATUS = 'QA HOLD' and QUANTITY > 0
                  and AVAILABLE_DATE = (case when DATEPART(WEEKDAY,getdate()) = 6 then convert(DATE,GETDATE()+3) else convert(DATE,GETDATE() +1) end)
                    """

    data = connection.GetSQLData(qa_query)

    # we want the QA at the end of inv list, so the skus in QA will be the last to be chose
    for line in data:
        inventory.append(INVObj(*line))  # add QA HOLD with inv

    return inventory


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


def satisfy_max_or_min(Wishes, Inventory, Parameters, satisfy_min=True, print_loads=False):
    """
    Attributes wishes wisely among p2p'S in Parameters list in order to satisfy their min or their max value

    :param Wishes: List of wishes (list of WishlistObj)
    :param Inventory: List of INVobj
    :param Parameters: List of Parameters
    :param satisfy_min: (bool) if false -> we want to satisfy the max
    :param print_loads: (bool) indicates if we plot each load or not
    """

    # We save a "trigger" integer value indicating if we want to satisfy min or max
    check_min = int(satisfy_min)  # Will be 1 if we want to satisfy min and 0 instead

    # For each parameters in Parameters list
    for param in Parameters:
        if len(param.LoadBuilder) < (check_min*param.LOADMIN + (1-check_min)*param.LOADMAX):

            # Initialization of empty list
            tempoOnLoad = []  # List to remember the INVobj that will be sent to the LoadBuilder
            invData = []      # List that will contain the data to build the frame that will be sent to the LoadBuilder

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
                        tempoOnLoad.append(wish)

                        # Here we set QTY and NBR_PER_CRATE to 1 because each line of the wishlist correspond to
                        # one crate and not one unit! Must be done this way to avoid having getting to many size_code
                        # in the returning list of the LoadBuilder
                        invData.append(wish.get_loadbuilder_input_line())

                        # We add the ranking of the wish in the ranking dictionary
                        if wish.SIZE_DIMENSIONS in ranking:
                            ranking[wish.SIZE_DIMENSIONS] += [wish.RANK]
                        else:
                            ranking[wish.SIZE_DIMENSIONS] = [wish.RANK]

            # Construction of the data frame which we'll send to the LoadBuilder of our parameters object (p2p)
            input_dataframe = loadbuilder_input_dataframe(invData)

            # Construction of loadings
            result = param.LoadBuilder.build(models_data=input_dataframe,
                                             max_load=(check_min*param.LOADMIN + (1-check_min)*param.LOADMAX),
                                             plot_load_done=print_loads,
                                             ranking=ranking)

            # Choice the wish items to put on loads
            for model, crate_type in result:
                found = False
                for OnLoad in tempoOnLoad:
                    if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0 and crate_type == OnLoad.CRATE_TYPE:
                        OnLoad.QUANTITY = 0
                        found = True
                        param.AssignedWish.append(OnLoad)
                        break
                if not found:
                    print('Error in Perfect Match: impossible result.\n')

            for wish in tempoOnLoad:  # If it is not on loads, give back inv
                if wish.QUANTITY > 0:
                    for inv in wish.INV_ITEMS:
                        inv.QUANTITY += 1
                    wish.INV_ITEMS = []


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


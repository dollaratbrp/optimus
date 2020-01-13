"""

Author : Olivier Lefebre

This file contained informations on WishlistObj class

Last update : 2020-01-06
By : Nicolas Raymond

"""

from LoadBuilder import LoadBuilder
from Import_Functions import *
DATAInclude = []


class WishListObj:
    def __init__(self, SDN, SINU, STN, PF, SP, DIV, MAT_NUM, SIZE, LENG, WIDTH, HEIGHT,
                 STACK, QTY, RANK, MANDATORY, OVERHANG, IsAdhoc=0):

        self.SALES_DOCUMENT_NUMBER = SDN
        self.SALES_ITEM_NUMBER = SINU
        self.SOLD_TO_NUMBER = STN
        self.POINT_FROM = PF
        self.SHIPPING_POINT = SP
        self.DIVISION = DIV
        self.MATERIAL_NUMBER = MAT_NUM
        self.SIZE_DIMENSIONS = SIZE
        self.LENGTH = LENG
        self.WIDTH = WIDTH
        self.HEIGHT = HEIGHT
        self.STACKABILITY = STACK
        self.QUANTITY = QTY
        self.RANK = RANK
        self.MANDATORY = MANDATORY
        self.OVERHANG = OVERHANG
        self.IsAdhoc=IsAdhoc

        # To keep track of inv origins
        self.INV_ITEMS = []
        self.ORIGINAL_QUANTITY = QTY

        # If the item is assigned to a Load in excel workbook
        self.Finished = False
        self.EndDate = None

    def lineToXlsx(self, dateToday):
        return [(self.SALES_DOCUMENT_NUMBER, self.SALES_ITEM_NUMBER, self.SOLD_TO_NUMBER, self.POINT_FROM,
                 self.SHIPPING_POINT, self.DIVISION, self.MATERIAL_NUMBER, self.SIZE_DIMENSIONS,
                 self.LENGTH, self.WIDTH, self.HEIGHT, self.STACKABILITY, self.OVERHANG, self.ORIGINAL_QUANTITY,
                 self.RANK, self.MANDATORY, self.EndDate, dateToday, self.IsAdhoc)]


class INVObj:
    def __init__(self, POINT, MATERIAL_NUMBER, QUANTITY, DATE, STATUS):
        self.POINT = POINT
        self.MATERIAL_NUMBER = MATERIAL_NUMBER
        self.QUANTITY = QUANTITY
        self.DATE = DATE
        self.STATUS = STATUS
        self.Future = not (weekdays(0) == DATE)  # if available date is not the same as today

        #To see if we took inv
        #self.ORIGINAL_QUANTITY = QUANTITY

        self.unused = 0  # count the number of skus to display on BOOKED_UNUSED worksheet

    def lineToXlsx(self):
        return [self.POINT, self.MATERIAL_NUMBER, self.QUANTITY, self.DATE, self.STATUS]


class Parameters:
    def __init__(self, POINT_FROM, POINT_TO, LOADMIN, LOADMAX, DRYBOX, FLATBED, TRANSIT, PRIORITY, days_to):
        self.POINT_FROM = POINT_FROM
        self.POINT_TO = POINT_TO
        self.LOADMIN = LOADMIN
        self.LOADMAX = LOADMAX
        self.DRYBOX = DRYBOX
        self.FLATBED = FLATBED
        self.PRIORITY = PRIORITY
        self.TRANSIT = TRANSIT
        self.days_to = days_to

        self.LoadBuilder = []
        self.new_LoadBuilder()
        self.AssignedWish = []

    def new_LoadBuilder(self):
        """"We reset the loadBuilder"""
        self.AssignedWish = []
        TrailerData = get_trailers_data(['DRYBOX', 'FLATBED'], [self.DRYBOX, self.FLATBED])
        self.LoadBuilder = LoadBuilder(TrailerData)


class Included_Inv:
    def __init__(self, Point_Source, Point_Include):
        self.source = Point_Source
        self.include = Point_Include


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
    for line in data:
        index_of_qty = category_list.index(line[0])
        line.insert(0, qty_list[index_of_qty])

    return pd.DataFrame(data=data, columns=columns)


def EquivalentPlantFrom(Point1, Point2):
    """" Point1 is shipping_point_from for inv, Point2 is shipping_point_from for wishlist
        Point1 is included in Point2                                                      """
    if Point1 == Point2:
        return True
    else:
        global DATAInclude
        for equiv in DATAInclude:
            if equiv.source == Point2 and equiv.include == Point1:
                return True
    return False


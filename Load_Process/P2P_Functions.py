from LoadBuilder import LoadBuilder
from Import_Functions import *
DATAInclude=[]
class WishListObj:
    def __init__(self,SDN,SINU,STN,PF,SP,DIV,MAT_NUM,SIZE,LENG,WIDTH,HEIGHT,STACK,QTY,RANK,MANDATORY,OVERHANG):
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
        self.OVERHANG=OVERHANG
        # To keep track of inv origins
        self.INV_ITEMS=[]

    def lineToXlsx(self):
        return [self.SALES_DOCUMENT_NUMBER, self.SALES_ITEM_NUMBER , self.SOLD_TO_NUMBER,self.POINT_FROM,self.SHIPPING_POINT,self.DIVISION,self.MATERIAL_NUMBER,
        self.SIZE_DIMENSIONS ,self.LENGTH ,self.WIDTH ,self.HEIGHT,self.STACKABILITY,self.QUANTITY,self.RANK,self.MANDATORY,SELF.OVERHANG]


class INVObj:
    def __init__(self,POINT,MATERIAL_NUMBER ,QUANTITY,DATE,STATUS):
        self.POINT = POINT
        self.MATERIAL_NUMBER=MATERIAL_NUMBER
        self.QUANTITY=QUANTITY
        self.DATE=DATE
        self.STATUS = STATUS
        #To see if we took inv
        self.ORIGINAL_QUANTITY = QUANTITY


    def lineToXlsx(self):
        return [self.POINT, self.MATERIAL_NUMBER,self.QUANTITY,self.DATE,self.STATUS]


class Parameters:
    def __init__(self,POINT_FROM,POINT_TO,LOADMIN,LOADMAX,DRYBOX,FLATBED,TRANSIT,PRIORITY):
        self.POINT_FROM = POINT_FROM
        self.POINT_TO=POINT_TO
        self.LOADMIN=LOADMIN
        self.LOADMAX=LOADMAX
        self.DRYBOX = DRYBOX
        self.FLATBED = FLATBED
        self.PRIORITY = PRIORITY
        self.TRANSIT = TRANSIT

        TrailerData = pd.DataFrame(
            data=[[self.FLATBED, 'FLATBED', 636, 102, 120, 1, 1],
                  [self.DRYBOX, 'DRYBOX',  628, 98, 120, 0, 1]],
            columns=['QTY', 'CATEGORY', 'LENGTH', 'WIDTH', 'HEIGHT', 'OVERHANG',
                     'PRIORITY_RANK'])
        self.LoadBuilder = LoadBuilder(TrailerData)

class Included_Inv:
    def __init__(self,Point_Source,Point_Include):
        self.source=Point_Source
        self.include = Point_Include

##Functions ##################################################################################

def EquivalentPlantFrom(Point1,Point2):
    """" Point1 is shipping_point_from for inv, Point2 is shipping_point_from for wishlist
        Point1 is included in Point2                                                      """
    if Point1== Point2:
        return True
    else:
        global DATAInclude
        for equiv in DATAInclude:
            if equiv.source == Point2 and equiv.include == Point1:
                return True
    return False


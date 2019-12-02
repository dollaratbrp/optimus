from LoadBuilder import LoadBuilder
from Import_Functions import *

class WishListObj:
    def __init__(self,SDN,SINU,STN,PF,SP,DIV,MAT_NUM,SIZE,LENG,WIDTH,HEIGHT,STACK,QTY,RANK,MANDATORY):
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
        self.On_Load = False
    def lineToXlsx(self):
        return [self.SALES_DOCUMENT_NUMBER, self.SALES_ITEM_NUMBER , self.SOLD_TO_NUMBER,self.POINT_FROM,self.SHIPPING_POINT,self.DIVISION,self.MATERIAL_NUMBER,
        self.SIZE_DIMENSIONS ,self.LENGTH ,self.WIDTH ,self.HEIGHT,self.STACKABILITY,self.QUANTITY,self.RANK,self.MANDATORY]


class INVObj:
    def __init__(self,POINT,MATERIAL_NUMBER ,QUANTITY,DATE,STATUS):
        self.POINT = POINT
        self.MATERIAL_NUMBER=MATERIAL_NUMBER
        self.QUANTITY=QUANTITY
        self.DATE=DATE
        self.STATUS = STATUS

    def lineToXlsx(self):
        return [self.POINT, self.MATERIAL_NUMBER,self.QUANTITY,self.DATE,self.STATUS]


class Parameters:
    def __init__(self,POINT_FROM,POINT_TO,LOADMIN,LOADMAX,DRYBOX,FLATBED,TRANSIT,PRIORITY,Loads_Made):
        self.POINT_FROM = POINT_FROM
        self.POINT_TO=POINT_TO
        self.LOADMIN=LOADMIN
        self.LOADMAX=LOADMAX
        self.DRYBOX = DRYBOX
        self.FLATBED = FLATBED
        self.PRIORITY = PRIORITY
        self.TRANSIT = TRANSIT
        self.Loads_Made=Loads_Made

        TrailerData = pd.DataFrame(
            data=[[self.FLATBED, 'FLATBED', self.POINT_FROM, self.POINT_TO, 636, 102, 120, 1, 1],
                  [self.DRYBOX, 'DRYBOX', self.POINT_FROM, self.POINT_TO, 628, 98, 120, 0, 1]],
            columns=['QTY', 'CATEGORY', 'PLANT_FROM', 'PLANT_TO', 'LENGTH', 'WIDTH', 'HEIGHT', 'OVERHANG',
                     'PRIORITY_RANK'])
        self.LoadBuilder = LoadBuilder(POINT_FROM, POINT_TO, [], TrailerData, weekdays(0))


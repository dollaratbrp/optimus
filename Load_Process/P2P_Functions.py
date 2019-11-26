class WishListObj:
    def __init__(self,SDN,SINU,STN,PF,SP,DIV,MAT_NUM,SIZE,LENG,WIDTH,HEIGHT,STACK,QTY,RANK,MANDATORY):
        self.SALES_DOCUMENT_NUMBER = SDN
        self.SALES_ITEM_NUMBER = SINU
        self.SOLD_TO_NUMBER = STN
        self.PLANT_FROM = PF
        self.SHIPPING_PLANT = SP
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
    def lineToXlsx(self):
        return [self.SALES_DOCUMENT_NUMBER, self.SALES_ITEM_NUMBER , self.SOLD_TO_NUMBER,self.PLANT_FROM,self.SHIPPING_PLANT,self.DIVISION,self.MATERIAL_NUMBER,
        self.SIZE_DIMENSIONS ,self.LENGTH ,self.WIDTH ,self.HEIGHT,self.STACKABILITY,self.QUANTITY,self.RANK,self.MANDATORY]


class INVObj:
    def __init__(self,PLANT,SHIPPING_POINT,MATERIAL_NUMBER ,DIVISION,INVENTORY):
        self.PLANT = PLANT
        self.SHIPPING_POINT=SHIPPING_POINT
        self.MATERIAL_NUMBER=MATERIAL_NUMBER
        self.DIVISION=DIVISION
        self.INVENTORY=INVENTORY


    def lineToXlsx(self):
        return [self.PLANT, self.SHIPPING_POINT,self.MATERIAL_NUMBER,self.DIVISION,self.INVENTORY]
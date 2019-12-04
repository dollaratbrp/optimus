#from Import_Functions import *
#from LoadBuilder import LoadBuilder
from ParametersBox import *
from P2P_Functions import *
import pandas as pd


if not OpenParameters():  # If user cancel request
    sys.exit()


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                                                     #MODIFICATIONS SECTION
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
dayTodayComplete= pd.datetime.now().replace(second=0, microsecond=0)
dayToday= weekdays(0)
            # os.path()
saveFolder='S:\Shared\Business_Planning\Personal\Lefebvre\S2\p2p\p2p_Project\p2p_Project\\' #Folder to save data
dest_filename = 'P2P_Summary_'+dayToday #Name of excel file with today's date
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                                                        #END OF MODIFICATIONS
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#


timeSinceLastCall('',FALSE)
#####################################################################################################################
                                                ###Get SQL DATA
# ####################################################################################################################

# #If SQL queries crash
downloaded = False
numberOfTry = 0
while (not downloaded and numberOfTry<3): # 3 trials, SQL Queries sometime crash for no reason
    numberOfTry+=1
    try:
        downloaded = True
####################

        ####################################################################################
                           ###  Email address Query
        ####################################################################################


        headerEmail='EMAIL_ADDRESS'
        SQLEmail = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS',headers=headerEmail)
        QueryEmail=""" SELECT distinct [EMAIL_ADDRESS]
         FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS]
        """
        #GET SQL DATA
        EmailList = [ sublist for sublist in SQLEmail.GetSQLData(QueryEmail) ]  #[ item for sublist in SQLEmail.GetSQLData(SQLEmail) for item in sublist]



        ####################################################################################
                           ###  Parameters Query
        ####################################################################################

        headerParams='POINT_FROM,POINT_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,TRANSIT,PRIORITY_ORDER,SKIP'
        SQLParams = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS',headers=headerParams)
        QueryParams=""" SELECT  [POINT_FROM]
              ,[POINT_TO]
              ,[LOAD_MIN]
              ,[LOAD_MAX]
              ,[DRYBOX]
              ,[FLATBED]
              ,[TRANSIT]
              ,[PRIORITY_ORDER]
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
          where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
          and SKIP = 0
          order by PRIORITY_ORDER
        """
        #GET SQL DATA
        DATAParams = [ Parameters(*sublist) for sublist in SQLParams.GetSQLData(QueryParams) ]  #[ item for sublist in SQLEmail.GetSQLData(SQLEmail) for item in sublist]


        ####################################################################################
                           ###  WishList Query
        ####################################################################################
        headerWishList = 'SALES_DOCUMENT_NUMBER,SALES_ITEM_NUMBER,SOLD_TO_NUMBER,POINT_FROM,SHIPPING_POINT,DIVISION,MATERIAL_NUMBER,Size_Dimensions,Lenght,Width,Height,stackability,Quantity,Priority_Rank,X_IF_MANDATORY'
        SQLWishList = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_2_PRIORITY_F_P2P',headers=headerWishList)
        QueryWishList= """SELECT  [SALES_DOCUMENT_NUMBER]
              ,[SALES_ITEM_NUMBER]
              ,[SOLD_TO_NUMBER]
              ,[POINT_FROM]
              ,[SHIPPING_POINT]
              ,[DIVISION]
              ,RTRIM([MATERIAL_NUMBER])
              ,RTRIM([Size_Dimensions])
              ,convert(int,CEILING([Length]))
              ,convert(int,CEILING([Width]))
              ,convert(int,CEILING([Height]))
              ,convert(int,[stackability])
              ,[Quantity]
              ,[Priority_Rank]
              ,[X_IF_MANDATORY]
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY]
          where Length<>0 and Width <> 0 and Height <> 0
          order by Priority_Rank
        """
        OriginalDATAWishList = SQLWishList.GetSQLData(QueryWishList)
        DATAWishList=[WishListObj(*obj) for obj in OriginalDATAWishList]
        # for obj in OriginalDATAWishList:
        #     DATAWishList.append(WishListObj(*obj))

        ####################################################################################
                           ###  INV Query
        ####################################################################################
        headerINV = ''#Not important here
        SQLINV = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_INVENTORY', headers=headerINV)
        QueryINV = """SELECT  [SHIPPING_POINT]
              ,[MATERIAL_NUMBER]
              ,case when [QUANTITY] <0 then 0 else convert(int,QUANTITY) end as qty
              ,[AVAILABLE_DATE]
              ,[STATUS]
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
          where status = 'INVENTORY'
        """
        OriginalDATAINV = SQLINV.GetSQLData(QueryINV)
        DATAINV = []
        for obj in OriginalDATAINV:
            DATAINV.append(INVObj(*obj))



        ####################################################################################
        ###  Included Shipping_point
        ####################################################################################
        headerInclude = ''  # Not important here
        SQLInclude = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_D_INCLUDED_INVENTORY', headers=headerInclude)
        QueryInclude = """select SHIPPING_POINT_SOURCE ,SHIPPING_POINT_INCLUDE
                            from OTD_1_P2P_D_INCLUDED_INVENTORY
            """
        OriginalDATAInclude = SQLInclude.GetSQLData(QueryInclude)
        #DATAInclude = [] # defined in P2P_Functions
        global DATAInclude
        for obj in OriginalDATAInclude:
            DATAInclude.append(Included_Inv(*obj))


        ####################################################################################
        ###  Look if all point_from + shipping_point are in parameters
        ####################################################################################

        SQLMissing = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PRIORITY', headers='')
        QueryMissing = """SELECT DISTINCT [POINT_FROM]
              ,[SHIPPING_POINT]
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY] 
          where CONCAT(POINT_FROM,SHIPPING_POINT) not in (
            select distinct CONCAT( [POINT_FROM],[POINT_TO])
            FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS] )
        """
        DATAMissing = SQLMissing.GetSQLData(QueryMissing)
    except:
        downloaded=False
        print('SQL Query failed')

#If SQL Queries failed
if not downloaded:
    try:
        print('failed')
        #send_email(EmailList, dest_filename, 'SQL QUERIES FAILED')
    except:
        pass
    sys.exit()


timeSinceLastCall('Get SQL DATA')


# if DATAMissing != []:
#     if not MissingP2PBox(DATAMissing):
#         sys.exit()

timeSinceLastCall('',False)
#####################################################################################################################
                                                ### Excel Workbook declaration
#####################################################################################################################


wb = Workbook()

### Summary
wsSummary = wb.active
wsSummary.title = "SUMMARY"



### Approved loads
wsApproved = wb.create_sheet("APPROVED")
wsApproved.append([''])
wsApprovedList = []


### Unbooked
wsUnbooked = wb.create_sheet("UNBOOKED")
wsUnbooked.append([''])
wsUnbookedList=[]



#####################################################################################################################
                                                ###Isolate perfect match
#####################################################################################################################
ListApprovedWish = []

for wish in DATAWishList:
    position = 0
    for Iteration in range( wish.QUANTITY ):
        for inv in range(position,len(DATAINV)):
            if  EquivalentPlantFrom(DATAINV[inv].POINT,wish.POINT_FROM) and wish.MATERIAL_NUMBER==DATAINV[inv].MATERIAL_NUMBER and DATAINV[inv].QUANTITY>0: #wish.POINT_FROM==inv.POINT and
                DATAINV[inv].QUANTITY-=1
                wish.INV_ITEMS.append(DATAINV[inv])
                position=inv
                break  # no need to look further


    if len(wish.INV_ITEMS) < wish.QUANTITY: #We give back taken inv
        for invToGiveBack in wish.INV_ITEMS:
            invToGiveBack.QUANTITY+=1
        wish.INV_ITEMS = []
    else:
        ListApprovedWish.append(wish)

### We don't need unbooked skus
for inv in DATAINV:
    if inv.QUANTITY>0:
        wsUnbooked.append(inv.lineToXlsx())





#####################################################################################################################
                                                ###Create Loads
#####################################################################################################################
for param in DATAParams:
    tempoOnLoad=[]
    columnsHead=['QTY','MODEL','LENGTH','WIDTH','HEIGHT','NBR_PER_CRATE','STACK_LIMIT','OVERHANG']
    invData = []#pd.DataFrame([],columns=['QTY','MODEL','PLANT_TO','LENGTH','WIDTH','HEIGHT','NBR_PER_CRATE','STACK_LIMIT','OVERHANG'])
    for wish in ListApprovedWish:
        if wish.POINT_FROM==param.POINT_FROM  and wish.SHIPPING_POINT==param.POINT_TO and wish.QUANTITY>0:
            tempoOnLoad.append(wish)
            invData.append([1, wish.SIZE_DIMENSIONS,wish.LENGTH,wish.WIDTH,wish.HEIGHT,1,wish.STACKABILITY,0])#quantity is one, one box for each line
    models_data = pd.DataFrame(data = invData, columns=columnsHead)


    result = param.LoadBuilder.build(models_data,param.LOADMAX, plot_load_done=False)

    for model in result:
        found = False
        for OnLoad in tempoOnLoad:
            if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY>0:
                OnLoad.QUANTITY =0
                found = True
                break
        if not found:
            print('Error in Perfect Match: impossible result.\n')


#####################################################################################################################
                                                ###Store unallocated units in inv pool
#####################################################################################################################
print('part2')
InventoryPool = []
for wish in ListApprovedWish:
    if wish.QUANTITY>0:
        for inv in wish.INV_ITEMS:
            InventoryPool.append(inv)


#####################################################################################################################
                                                ###Make minimum number of loads for each P2P
#####################################################################################################################
for param in DATAParams:
    if len(param.LoadBuilder) < param.LOADMIN:
        tempoOnLoad = []
        columnsHead = ['QTY', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']
        invData = []
        for wish in DATAWishList:
            if wish.POINT_FROM==param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO and wish.QUANTITY>0:
                position =0
                for Iteration in range(wish.QUANTITY):
                    for inv in range(position, len(InventoryPool)):
                        if EquivalentPlantFrom(InventoryPool[inv].POINT,wish.POINT_FROM) and InventoryPool[inv].MATERIAL_NUMBER==wish.MATERIAL_NUMBER and InventoryPool[inv].QUANTITY >0:
                            InventoryPool[inv].QUANTITY -= 1
                            wish.INV_ITEMS.append(InventoryPool[inv])
                            position = inv
                            break # no need to look further
                if len(wish.INV_ITEMS) < wish.QUANTITY:  # We give back taken inv
                    for invToGiveBack in wish.INV_ITEMS:
                        invToGiveBack.QUANTITY += 1
                    wish.INV_ITEMS=[]
                else:
                    tempoOnLoad.append(wish)
                    invData.append([1, wish.SIZE_DIMENSIONS, wish.LENGTH, wish.WIDTH, wish.HEIGHT, 1, wish.STACKABILITY, 0])

        models_data = pd.DataFrame(data=invData, columns=columnsHead)
        print(models_data)
        result = param.LoadBuilder.build(models_data, param.LOADMIN - len(param.LoadBuilder), plot_load_done=False)
        print(result)

#####################################################################################################################
                                                ###Test to save data
#####################################################################################################################



reference = [savexlsxFile(wb, saveFolder, dest_filename)]
#send_email(EmailList, dest_filename, 'generalErrors?', reference)
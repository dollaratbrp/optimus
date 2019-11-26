from Import_Functions import *
from ParametersBox import *
from P2P_Functions import *

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
# while (not downloaded and numberOfTry<3): # 3 trials, SQL Queries sometime crash for no reason
##########
numberOfTry+=1
downloaded = True
#    try:
####################

        ####################################################################################
                           ###  Email address Query
        ####################################################################################
        ###Use class

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

headerParams='PLANT_FROM,PLANT_TO,LOAD_MIN,LOAD_MAX,DRYBOX,FLATBED,TRANSIT,PRIORITY_ORDER,SKIP'
SQLParams = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS',headers=headerParams)
QueryParams=""" SELECT  [PLANT_FROM]
      ,[PLANT_TO]
      ,[LOAD_MIN]
      ,[LOAD_MAX]
	  ,[DRYBOX]
      ,[FLATBED]
      ,[TRANSIT]
	  ,[PRIORITY_ORDER]
      ,[SKIP]
      --,[IMPORT_DATE]
  FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
  where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
  and SKIP = 0
  order by PRIORITY_ORDER
"""
#GET SQL DATA
DATAParams = [ sublist for sublist in SQLParams.GetSQLData(QueryParams) ]  #[ item for sublist in SQLEmail.GetSQLData(SQLEmail) for item in sublist]


        ####################################################################################
                           ###  WishList Query
        ####################################################################################
headerWishList = 'SALES_DOCUMENT_NUMBER,SALES_ITEM_NUMBER,SOLD_TO_NUMBER,PLANT_FROM,SHIPPING_PLANT,DIVISION,MATERIAL_NUMBER,Size_Dimensions,Lenght,Width,Height,stackability,Quantity,Priority_Rank,X_IF_MANDATORY'
SQLWishList = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_2_PRIORITY_F_P2P',headers=headerWishList)
QueryWishList= """SELECT  [SALES_DOCUMENT_NUMBER]
      ,[SALES_ITEM_NUMBER]
      ,[SOLD_TO_NUMBER]
      ,[PLANT_FROM]
      ,[SHIPPING_PLANT]
      ,[DIVISION]
      ,RTRIM([MATERIAL_NUMBER])
      ,RTRIM([Size_Dimensions])
      ,[Lenght]
      ,[Width]
      ,[Height]
      ,[stackability]
      ,[Quantity]
      ,[Priority_Rank]
      ,[X_IF_MANDATORY]
  FROM [Business_Planning].[dbo].[OTD_2_PRIORITY_F_P2P]
  order by Priority_Rank
"""
OriginalDATAWishList = SQLWishList.GetSQLData(QueryWishList)
DATAWishList=[]
for obj in OriginalDATAWishList:
    DATAWishList.append(WishListObj(*obj))

        ####################################################################################
                           ###  INV Query
        ####################################################################################
headerINV = ''
SQLINV = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_SHIPPING_F_INVENTORY_VISIBILITY', headers=headerINV)
QueryINV = """SELECT [PLANT]
  ,[SHIPPING_POINT]
  ,[MATERIAL_NUMBER]
  ,[DIVISION]
  ,[INVENTORY]
FROM [Business_Planning].[dbo].[OTD_1_SHIPPING_F_INVENTORY_VISIBILITY]
"""
OriginalDATAINV = SQLINV.GetSQLData(QueryINV)
DATAINV = []
for obj in OriginalDATAINV:
    DATAINV.append(INVObj(*obj))


        ####################################################################################
                           ###  Parameters Query
        ####################################################################################

#    except:
#        downloaded=False
#        print('SQL Query failed')  

#If SQL Queries failed
if not downloaded:
    try:
        send_email(EmailList, dest_filename, 'SQL QUERIES FAILED')
    except:
        pass
    sys.exit()


timeSinceLastCall('Get SQL DATA')

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
                                                ### Lists declaration
#####################################################################################################################







#####################################################################################################################
                                                ###Isolate perfect match
#####################################################################################################################
ListApproved = []
ListUnbooked = []

for wish in DATAWishList:
    found = False
    for inv in DATAINV:
        if wish.PLANT_FROM==inv.PLANT and wish.MATERIAL_NUMBER==inv.MATERIAL_NUMBER and inv.INVENTORY>0:
            inv.INVENTORY-=1
            ListApproved.append(wish)
            found = True
            break #no need to look further
    if not found:
        ListUnbooked.append(wish)

### We don't need unbooked skus
for wish in ListUnbooked:
    wsUnbooked.append(wish.lineToXlsx())



#####################################################################################################################
                                                ###Test to save data
#####################################################################################################################



#for obj in DATAWishList:
 #   wsUnbooked.append(obj.lineToXlsx())
for obj in DATAINV:
    wsApproved.append(obj.lineToXlsx())

reference = [savexlsxFile(wb, saveFolder, dest_filename)]
#send_email(EmailList, dest_filename, 'generalErrors?', reference)
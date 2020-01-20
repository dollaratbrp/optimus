# P2P Forecast for next 8 weeks

### Used SQL Table and actions:
#[Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS] : Get data, alter, delete, send
#[Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS] : Get data, alter, delete, send
#[Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY] : Get data
#[Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY] : Get data
#OTD_1_P2P_D_INCLUDED_INVENTORY : Get data
#OTD_1_P2P_F_FORECAST_LOADS : Send data
#OTD_1_P2P_F_FORECAST_PRIORITY : Send data

### Important path
# None

from ParametersBox import *
from P2P_Functions import *

AutomaticRun = False # set to True to automate code

if not AutomaticRun:
    if not OpenParameters('FORECAST'):  # If user cancel request
        sys.exit()
    IsAdhoc = 1  # 1 for True, 0 for False (if automatic)
else:
    IsAdhoc = 0  # 1 for True, 0 for False (if automatic)



# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                                                     #MODIFICATIONS SECTION
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
dayTodayComplete= pd.datetime.now().replace(second=0, microsecond=0)
dayToday= weekdays(0)
printLoads = False

dest_filename = 'P2P_Forecast_'+dayToday # Email subject with today's date
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                                                        #END OF MODIFICATIONS
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
GeneralErrors='' #General errors met during execution


#####################################################################################################################
                                                ###Get SQL DATA
# ####################################################################################################################

# #If SQL queries crash
downloaded = False
numberOfTry = 0
while not downloaded and numberOfTry<3: # 3 trials, SQL Queries sometime crash for no reason
    numberOfTry+=1
    try:
        downloaded = True
        ####################################################################################
                           ###  Email address Query
        ####################################################################################
        headerEmail='EMAIL_ADDRESS'
        SQLEmail = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS',headers=headerEmail)
        QueryEmail=""" SELECT distinct [EMAIL_ADDRESS]
         FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS_EMAIL_ADDRESS]
         WHERE PROJECT = 'FORECAST'
        """
        #GET SQL DATA
        EmailList = [ item for sublist in SQLEmail.GetSQLData(QueryEmail) for item in sublist]#[ sublist for sublist in SQLEmail.GetSQLData(QueryEmail) ]  #[ item for sublist in SQLEmail.GetSQLData(SQLEmail) for item in sublist]

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
              ,DAYS_TO
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS]
          where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS])
          and SKIP = 0
          order by PRIORITY_ORDER
        """
        #GET SQL DATA
        DATAParams = [ Parameters(*sublist) for sublist in SQLParams.GetSQLData(QueryParams) ]  #[ item for sublist in SQLEmail.GetSQLData(SQLEmail) for item in sublist]

        ####################################################################################
                           ### Distinct date
        ####################################################################################
        ##we use the dates to loop through time
        SQLINV = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_INVENTORY',
                                  headers='')
        QueryDATE = """ SELECT distinct convert(date,[AVAILABLE_DATE]) as [AVAILABLE_DATE]
                  FROM [Business_Planning].[dbo].[OTD_1_P2P_F_INVENTORY]
                  where [AVAILABLE_DATE] between convert(date,getdate()+1) and convert(date,getdate() + 7*8)--we want the next 8 weeks
                  order by [AVAILABLE_DATE]
                """

        DATADATE = [single for sublist in SQLINV.GetSQLData(QueryDATE) for single in sublist]

        ####################################################################################
                           ### First Inv Query
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

        DATAINV = [INVObj(*obj) for obj in SQLINV.GetSQLData(QueryInv)]  # [sublist for sublist in SQLINV.GetSQLData(QueryDATE)]

        ####################################################################################
        ### Prod & QA Query
        ####################################################################################
        ##Filter based on if production is late, to make later
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

        DATAProd = [INVObj(*obj) for obj in
                   SQLINV.GetSQLData(QueryProd)]  # [sublist for sublist in SQLINV.GetSQLData(QueryDATE)]


        ####################################################################################
                           ###  WishList Query
        ####################################################################################

        SQLWishList = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_2_PRIORITY_F_P2P',headers='')
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
              ,[OVERHANG]
              ,"""+str(IsAdhoc)+""" as IsAdhoc
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY]
          where [POINT_FROM] <>[SHIPPING_POINT] and Length<>0 and Width <> 0 and Height <> 0
          and concat([POINT_FROM],[SHIPPING_POINT]) in (
		  SELECT distinct concat( [POINT_FROM] ,[POINT_TO])
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS]
          where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_FORECAST_PARAMETERS])
          and SKIP = 0 )
          order by [X_IF_MANDATORY] desc, Priority_Rank
        """

        OriginalDATAWishList = SQLWishList.GetSQLData(QueryWishList)
        DATAWishList=[WishListObj(*obj) for obj in OriginalDATAWishList]

        ####################################################################################
        ###  Included Shipping_point
        ####################################################################################

        SQLInclude = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_D_INCLUDED_INVENTORY', headers='')
        QueryInclude = """select SHIPPING_POINT_SOURCE ,SHIPPING_POINT_INCLUDE
                            from OTD_1_P2P_D_INCLUDED_INVENTORY
            """
        OriginalDATAInclude = SQLInclude.GetSQLData(QueryInclude)
        #DATAInclude = [] # defined in P2P_Functions
        global DATAInclude
        for obj in OriginalDATAInclude:
            DATAInclude.append(Included_Inv(*obj))

    except:
        downloaded = False
        print('SQL Query failed')



#If SQL Queries failed
if not downloaded:
    try:
        print('failed')
        send_email(EmailList, dest_filename, 'SQL QUERIES FAILED')
    except:
        pass
    sys.exit()


#####################################################################################################################
                                                ### SQL tables declaration
#####################################################################################################################
headerLoads = 'POINT_FROM,SHIPPING_POINT,QUANTITY,DATE,IMPORT_DATE,IS_ADHOC'
SQLLoads = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_FORECAST_LOADS', headers=headerLoads)

headerPRIORITY = 'SALES_DOCUMENT_NUMBER,SALES_ITEM_NUMBER,SOLD_TO_NUMBER,POINT_FROM,SHIPPING_POINT,DIVISION,MATERIAL_NUMBER,SIZE_DIMENSIONS,LENGTH,WIDTH,HEIGHT,STACKABILITY,OVERHANG,QUANTITY,PRIORITY_RANK,X_IF_MANDATORY,PROJECTED_DATE,IMPORT_DATE,IS_ADHOC'
SQLPRIORITY = SQLConnection('CAVLSQLPD2\pbi2', 'Business_Planning', 'OTD_1_P2P_F_FORECAST_PRIORITY', headers=headerPRIORITY)

#####################################################################################################################
                                                ### Main loop for everyday
#####################################################################################################################

#we add today's prod and QA to inventory, but we don't create load for today (we make them for tomorrow or later)
for prod in DATAProd:
    if prod.DATE > weekdays(0): #Ordered by date, so no need to continue
        break
    elif prod.DATE == weekdays(0) and prod.QUANTITY>0 and prod.STATUS== 'QA HOLD':
        for inv in DATAINV:
            if inv.POINT == prod.POINT and inv.MATERIAL_NUMBER == prod.MATERIAL_NUMBER:
                inv.QUANTITY+=prod.QUANTITY
                prod.QUANTITY=0
                break #we found the good inv

# Inventory is now updated.



for date in DATADATE:
    timeSinceLastCall('Loop')
    #we add today's prod and QA to inventory
    for prod in DATAProd:
        if prod.DATE > date: #Ordered by date, so no need to continue
            break
        elif prod.DATE == date and prod.QUANTITY>0:
            for inv in DATAINV:
                if inv.POINT == prod.POINT and inv.MATERIAL_NUMBER == prod.MATERIAL_NUMBER:
                    inv.QUANTITY+=prod.QUANTITY
                    prod.QUANTITY=0
                    break #we found the good inv

    # Inventory is now updated.

    # Now we create new loadBuilders
    for param in DATAParams:
        param.new_LoadBuilder()

    # Isolate perfect match

    ListApprovedWish = []

    for wish in DATAWishList:
        if wish.QUANTITY>0:
            position = 0 #to not loop again through the complete list
            for Iteration in range(wish.QUANTITY):
                for It, inv in enumerate(DATAINV[position::]):
                    if EquivalentPlantFrom(inv.POINT,
                                           wish.POINT_FROM) and wish.MATERIAL_NUMBER == inv.MATERIAL_NUMBER and inv.QUANTITY > 0:
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
        # C'est moins efficace d'utiliser la méthode ci-dessous (problème de priorité de la wishList)
    # for inv in DATAINV:
    #     if inv.QUANTITY>0:
    #         for wish in DATAWishList:
    #             if 0 == len(wish.INV_ITEMS) < wish.QUANTITY <= inv.QUANTITY and EquivalentPlantFrom(inv.POINT,wish.POINT_FROM) and wish.MATERIAL_NUMBER == inv.MATERIAL_NUMBER:
    #                 for It in range(wish.QUANTITY):
    #                     wish.INV_ITEMS.append(inv)
    #                     inv.QUANTITY-=1
    #                 ListApprovedWish.append(wish)
    #             if inv.QUANTITY == 0:#we gave all available inv
    #                 break
    #
    # # All available inventory is now assigned to a wish, not already satisfied
    # timeSinceLastCall('perfect match')

    #We now create loads with those perfect match
    for param in DATAParams:
        tempoOnLoad = []
        columnsHead = ['QTY', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']
        invData = []
        for wish in ListApprovedWish:
            if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO and wish.QUANTITY > 0:
                tempoOnLoad.append(wish)
                invData.append([1, wish.SIZE_DIMENSIONS, wish.LENGTH, wish.WIDTH, wish.HEIGHT, 1, wish.STACKABILITY,
                                wish.OVERHANG])  # quantity is one, one box for each line
        models_data = pd.DataFrame(data=invData, columns=columnsHead)
        result = param.LoadBuilder.build(models_data, param.LOADMAX, plot_load_done=printLoads)
        for model in result:
            found = False
            for OnLoad in tempoOnLoad:
                if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0: #we send allocated wish units to sql
                    OnLoad.QUANTITY = 0
                    found = True
                    param.AssignedWish.append(OnLoad)
                    OnLoad.EndDate = weekdays(max(0, param.days_to-1),officialDay = date)
                    SQLPRIORITY.sendToSQL(OnLoad.lineToXlsx(dayTodayComplete))
                    DATAWishList.remove(OnLoad)
                    break
            if not found:
                GeneralErrors+='Error in Perfect Match: impossible result.\n'

    # Store unallocated units in inv pool
    for wish in ListApprovedWish:
        if wish.QUANTITY > 0:
            for inv in wish.INV_ITEMS:
                inv.QUANTITY += 1
            wish.INV_ITEMS = []

    # Try to Make the minimum number of loads for each P2P

    for param in DATAParams:
        if len(param.LoadBuilder) < param.LOADMIN: #If the minimum isn't reached
            tempoOnLoad = []
            columnsHead = ['QTY', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']
            invData = []
            # create data table
            for wish in DATAWishList:
                if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO and wish.QUANTITY > 0:
                    position = 0
                    for Iteration in range(wish.QUANTITY):
                        for It, inv in enumerate(DATAINV[position::]):
                            if EquivalentPlantFrom(inv.POINT,
                                                   wish.POINT_FROM) and inv.MATERIAL_NUMBER == wish.MATERIAL_NUMBER and inv.QUANTITY > 0:
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
                            [1, wish.SIZE_DIMENSIONS, wish.LENGTH, wish.WIDTH, wish.HEIGHT, 1, wish.STACKABILITY,
                             wish.OVERHANG])
            # Create loads
            models_data = pd.DataFrame(data=invData, columns=columnsHead)
            result = param.LoadBuilder.build(models_data, param.LOADMIN - len(param.LoadBuilder),
                                             plot_load_done=printLoads)
            # Choose wish items to put on loads
            for model in result:
                found = False
                for OnLoad in tempoOnLoad:
                    if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0:#we send allocated wish units to sql
                        OnLoad.QUANTITY = 0
                        found = True
                        param.AssignedWish.append(OnLoad)
                        OnLoad.EndDate = weekdays(max(0, param.days_to - 1), officialDay=date)
                        SQLPRIORITY.sendToSQL(OnLoad.lineToXlsx(dayTodayComplete))
                        DATAWishList.remove(OnLoad)
                        break
                if not found:
                    GeneralErrors+='Error in min section: impossible result.\n'
            for wish in tempoOnLoad:  # If it is not on loads, give back inv
                if wish.QUANTITY > 0:
                    for inv in wish.INV_ITEMS:
                        inv.QUANTITY += 1
                    wish.INV_ITEMS = []

    # Try to Make the maximum number of loads for each P2P

    for param in DATAParams:
        if len(param.LoadBuilder) < param.LOADMAX: #if we haven't reached the max number of loads
            tempoOnLoad = []
            columnsHead = ['QTY', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']
            invData = []

            # Create data table
            for wish in DATAWishList:
                if wish.POINT_FROM == param.POINT_FROM and wish.SHIPPING_POINT == param.POINT_TO and wish.QUANTITY > 0:
                    position = 0
                    for Iteration in range(wish.QUANTITY):
                        for It, inv in enumerate(DATAINV[position::]):
                            if EquivalentPlantFrom(inv.POINT,
                                                   wish.POINT_FROM) and inv.MATERIAL_NUMBER == wish.MATERIAL_NUMBER and inv.QUANTITY > 0 :
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
                            [1, wish.SIZE_DIMENSIONS, wish.LENGTH, wish.WIDTH, wish.HEIGHT, 1, wish.STACKABILITY,
                             wish.OVERHANG])

            models_data = pd.DataFrame(data=invData, columns=columnsHead)
            # Create loads
            result = param.LoadBuilder.build(models_data, param.LOADMAX, plot_load_done=printLoads)
            # choose wish items to put on loads
            for model in result:
                found = False
                for OnLoad in tempoOnLoad:
                    if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0:#we send allocated wish units to sql
                        OnLoad.QUANTITY = 0
                        found = True
                        param.AssignedWish.append(OnLoad)
                        OnLoad.EndDate = weekdays(max(0, param.days_to - 1), officialDay=date)
                        SQLPRIORITY.sendToSQL(OnLoad.lineToXlsx(dayTodayComplete))
                        DATAWishList.remove(OnLoad)
                        break
                if not found:
                    GeneralErrors+='Error in max section: impossible result.\n'
            for wish in tempoOnLoad:  # If it is not on loads, give back inv
                if wish.QUANTITY > 0:
                    for inv in wish.INV_ITEMS:
                        inv.QUANTITY += 1
                    wish.INV_ITEMS = []


    # we save the loads created for that day
    for param in DATAParams:
        SQLLoads.sendToSQL([(param.POINT_FROM,param.POINT_TO,len(param.LoadBuilder),
                             weekdays(max(0, param.days_to-1),officialDay = date),dayTodayComplete ,IsAdhoc)])
        #print(param.POINT_FROM,param.POINT_TO,len(param.LoadBuilder), len(param.AssignedWish))

# Loads were made for the next 8 weeks, now we send not assigned wishes to SQL

for wish in DATAWishList:
    if wish.QUANTITY != wish.ORIGINAL_QUANTITY:
        GeneralErrors+='Error : this unit was not assigned and has a different quantity; \n {0}'.format(wish.lineToXlsx())
    SQLPRIORITY.sendToSQL(wish.lineToXlsx(dayTodayComplete))



# Finally, we send an email
send_email(EmailList, dest_filename, 'P2P Forecast is now updated.\n'+GeneralErrors)
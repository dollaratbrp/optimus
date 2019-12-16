
from ParametersBox import *
from P2P_Functions import *

# if not OpenParameters('FORECAST'):  # If user cancel request
#     sys.exit()



# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                                                     #MODIFICATIONS SECTION
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
dayTodayComplete= pd.datetime.now().replace(second=0, microsecond=0)
dayToday= weekdays(0)
printLoads = False

saveFolder='S:\Shared\Business_Planning\Personal\Lefebvre\S2\p2p\p2p_Project\p2p_Project\\' #Folder to save data
dest_filename = 'P2P_Forecast_'+dayToday #Name of excel file with today's date
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                                                        #END OF MODIFICATIONS
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------#



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
              ,DAYS_TO
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
          where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
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
                  where [AVAILABLE_DATE] between convert(date,getdate()) and convert(date,getdate() + 7*8)--we want the next 8 weeks
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
                          and AVAILABLE_DATE between convert(date,getdate()) and convert(date, getdate() + 10*7)
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
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PRIORITY_WITHOUT_INVENTORY]
          where [POINT_FROM] <>[SHIPPING_POINT] and Length<>0 and Width <> 0 and Height <> 0
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
        #send_email(EmailList, dest_filename, 'SQL QUERIES FAILED')
    except:
        pass
    sys.exit()


#####################################################################################################################
                                                ### Main loop for everyday
#####################################################################################################################
for date in DATADATE:
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
    # All available inventory is now assigned to a wish, not already satisfied


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
                if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0:
                    OnLoad.QUANTITY = 0
                    found = True
                    param.AssignedWish.append(OnLoad)
                    break
            if not found:
                print('Error in Perfect Match: impossible result.\n')

    # Store unallocated units in inv pool
    for wish in ListApprovedWish:
        if wish.QUANTITY > 0:
            for inv in wish.INV_ITEMS:
                inv.QUANTITY += 1
            wish.INV_ITEMS = []

    # Try to Make the minimum number of loads for each P2P

    for param in DATAParams:
        if len(param.LoadBuilder) < param.LOADMIN:
            tempoOnLoad = []
            columnsHead = ['QTY', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']
            invData = []
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

            models_data = pd.DataFrame(data=invData, columns=columnsHead)
            result = param.LoadBuilder.build(models_data, param.LOADMIN - len(param.LoadBuilder),
                                             plot_load_done=printLoads)
            for model in result:
                found = False
                for OnLoad in tempoOnLoad:
                    if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0:
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

    # Try to Make the maximum number of loads for each P2P

    for param in DATAParams:
        if len(param.LoadBuilder) < param.LOADMAX:
            tempoOnLoad = []
            columnsHead = ['QTY', 'MODEL', 'LENGTH', 'WIDTH', 'HEIGHT', 'NBR_PER_CRATE', 'STACK_LIMIT', 'OVERHANG']
            invData = []
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

            result = param.LoadBuilder.build(models_data, param.LOADMAX, plot_load_done=printLoads)

            for model in result:
                found = False
                for OnLoad in tempoOnLoad:
                    if OnLoad.SIZE_DIMENSIONS == model and OnLoad.QUANTITY > 0:
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


    # we save the loads created for that day
    print(date)
    for param in DATAParams:
        print(param.POINT_FROM,param.POINT_TO,len(param.LoadBuilder), len(param.AssignedWish))














from ParametersBox import *
from P2P_Functions import *

if not OpenParameters('FORECAST'):  # If user cancel request
    sys.exit()



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
          FROM [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS]
          where IMPORT_DATE = (select max(IMPORT_DATE) from [Business_Planning].[dbo].[OTD_1_P2P_F_PARAMETERS])
          and SKIP = 0
          order by PRIORITY_ORDER
        """
        #GET SQL DATA
        DATAParams = [ Parameters(*sublist) for sublist in SQLParams.GetSQLData(QueryParams) ]  #[ item for sublist in SQLEmail.GetSQLData(SQLEmail) for item in sublist]





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



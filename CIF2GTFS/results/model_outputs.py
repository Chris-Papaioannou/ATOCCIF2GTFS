import sys
import os

sys.path.append(os.path.join(os.path.dirname(__file__), "src"))


import pandas as pd
import numpy as np
import datetime
import tempfile
from pathlib import Path
import sqlite3


#* Set parquet compression type here
#* From: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_parquet.html:
#* Name of the compression to use. Use None for no compression. Supported options: ‘snappy’, ‘gzip’, ‘brotli’, ‘lz4’, ‘zstd’.
parquetCompression = 'snappy'

pathLegCols = {"PATH\PUTRELATION\ODPAIR\FROMZONE\CODE":str,
               "PATH\PUTRELATION\ODPAIR\TOZONE\CODE":str,
               "PATHINDEX":int,
               "PATHLEGINDEX":int,
               "PATH\COUNT:PUTPATHLEGSWITHWALK":int,
               "ODTRIPS":float,
               "FROMSTOPPOINTNO":float,
               "TOSTOPPOINTNO":float,
               "TIMEPROFILEKEYSTRING":str,
               "TIME":int,
               "WAITTIME":int,
               "STARTVEHJOURNEYITEM\VEHJOURNEY\ATOC":str,
               "STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINUID":str,
               "STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINSERVICECODE":str,
               "DEPTIME":str}

stopPointCols = {"NO":int,
                 "CODE":str,
                 "NAME":str,
                 "STOPAREA\STOP\CRS":str}


def vlog(msg, PRIO=20480):
    Visum.Log(PRIO, msg)


def getPathLegs(cols, tempPath, flowBundle, quitVisum):
    global Visum

    PuTPathLegList = Visum.Workbench.Lists.CreatePuTPathLegList
    for col in cols.keys():
        PuTPathLegList.AddColumn(col)

    PuTPathLegList.Show()

    if flowBundle:
        routes = 1
    else:
        routes = 0
    PuTPathLegList.SetObjects(0,"X",routes, ListFormat=2)

    timecode = datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')
    PuTPathLegList.SaveToSQLiteDatabase(f"{tempPath}\\PuTPathLegs_{timecode}.sqlite3", "PathLegs")

    if quitVisum:
        Visum = None

    SQL_Query = 'SELECT PATHINDEX, PATHLEGINDEX, "PATH\PUTRELATION\ODPAIR\FROMZONE\CODE", "PATH\PUTRELATION\ODPAIR\TOZONE\CODE", "PATH\COUNT:PUTPATHLEGSWITHWALK", ODTRIPS, FROMSTOPPOINTNO, TOSTOPPOINTNO, TIMEPROFILEKEYSTRING, TIME, WAITTIME, "STARTVEHJOURNEYITEM\VEHJOURNEY\ATOC", "STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINUID", "STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINSERVICECODE", DEPTIME FROM PathLegs WHERE TIMEPROFILEKEYSTRING NOT IN ("Origin connector", "Destination connector", "PuT Aux PuTAux") AND PATHLEGINDEX != 0'

    con = sqlite3.connect(f"{tempPath}\\PuTPathLegs_{timecode}.sqlite3") 
    dfPathLegs = pd.read_sql_query(SQL_Query, con, dtype=cols)# , chunksize=10000

    dfPathLegs.rename({"PATH\PUTRELATION\ODPAIR\FROMZONE\CODE":'OrigCRS', 'PATH\PUTRELATION\ODPAIR\TOZONE\CODE':'DestCRS', 'STARTVEHJOURNEYITEM\VEHJOURNEY\ATOC':'ATOC', 'STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINUID':'TrainUID', 'STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINSERVICECODE':'TrainServiceCode', "PATH\COUNT:PUTPATHLEGSWITHWALK":"NumLegs" }, axis=1, inplace=True, errors='ignore')

    for col in ['FROMSTOPPOINTNO', 'TOSTOPPOINTNO']:
        dfPathLegs[col].fillna(-1, inplace=True)
        dfPathLegs[col] = dfPathLegs[col].astype(int)


    for col in ['TrainUID', 'TrainServiceCode', 'ATOC']:
        dfPathLegs[col].fillna("", inplace=True)

    for col in ['TIME', 'WAITTIME']:
        dfPathLegs[col] = dfPathLegs[col]/60

    dfPathLegs['DEPTIME'] = pd.to_datetime(dfPathLegs.DEPTIME)
    dfPathLegs['DEPTIME']

    dfPathLegs['PATHLEGINDEX'] = dfPathLegs.PATHLEGINDEX.astype(int)
    dfPathLegs['NumLegs'] = dfPathLegs.NumLegs.astype(int)
    dfPathLegs['MovementType'] = np.where(dfPathLegs.PATHLEGINDEX == 2,
                                          'FirstRailLeg',
                                          np.where(dfPathLegs.PATHLEGINDEX == dfPathLegs.NumLegs-1,
                                                   'LastRailLeg',
                                                   np.where(dfPathLegs.TIMEPROFILEKEYSTRING == "Transfer",
                                                            np.where(dfPathLegs.TOSTOPPOINTNO==-1,
                                                                     "TubeTransfer",
                                                                     np.where(dfPathLegs.FROMSTOPPOINTNO==-1,
                                                                             "TransferFromTube",
                                                                             "RailTransfer"
                                                                            )
                                                                    ),
                                                            "IntermediateRailLeg"
                                                        )
                                                )
                                        )
                                         
    
    dfPathLegs.drop(['NumLegs', 'TIMEPROFILEKEYSTRING'], axis=1, inplace=True)


    dfPathLegs['TOSTOPPOINTNO'] = np.where((dfPathLegs.MovementType=='TubeTransfer') & (dfPathLegs.PATHINDEX == dfPathLegs.PATHINDEX.shift(-1)),
                                            dfPathLegs.TOSTOPPOINTNO.shift(-1),
                                            dfPathLegs.TOSTOPPOINTNO)
    
    dfPathLegs['ATOC'] = np.where(dfPathLegs.MovementType == 'TubeTransfer',
                                  'Tube',
                                  dfPathLegs.ATOC)

    for col in ['TIME', 'WAITTIME']:
        dfPathLegs[col] = np.where((dfPathLegs.MovementType=='TubeTransfer') &(dfPathLegs.PATHINDEX == dfPathLegs.PATHINDEX.shift(-1)),
                                    dfPathLegs[col]+dfPathLegs[col].shift(-1),
                                    dfPathLegs[col])
    
    dfPathLegs = dfPathLegs.loc[(dfPathLegs.MovementType!='TransferFromTube')|(dfPathLegs.PATHLEGINDEX==3)].copy()

    dfPathLegs.drop(['PATHLEGINDEX'], axis=1, inplace=True)

    dfPathLegs['MovementType'] = np.where(dfPathLegs.MovementType=='TransferFromTube',
                                          'TubeTransfer',
                                          dfPathLegs.MovementType)

    con = None
    
    return dfPathLegs, timecode


def getStopPoints(cols):

    global Visum

    SPList = Visum.Workbench.Lists.CreateStopPointBaseList
    for col in cols:
        SPList.AddColumn(col)
    dfStopPoints = pd.DataFrame(SPList.SaveToArray(), columns=cols)

    for col in dfStopPoints.columns.values:
        dfStopPoints[col] = dfStopPoints[col].astype(cols[col])

    dfStopPoints.rename({"STOPAREA\STOP\CRS":'CRS'}, axis=1, inplace=True, errors='ignore')

    return dfStopPoints

def runFlowBundle(CRS):
    FlowBundle = Visum.Net.DemandSegments.ItemByKey("X").FlowBundle
    AllStopsdict = dict(Visum.Net.Stops.GetMultipleAttributes(['CRS', 'NO']))

    stopNumber = int(AllStopsdict[CRS])
    stop = Visum.Net.Stops.ItemByKey(stopNumber)

    NetElementContainer = Visum.CreateNetElements()
    NetElementContainer.Add(stop)

    ActivityTypeSet1 = FlowBundle.CreateActivityTypeSet()
    ActivityTypeSet1.Add(4)
    ActivityTypeSet1.Add(8)
    ActivityTypeSet1.Add(16)

    FlowBundle.CreateConditionWithRestrictedSupply(stop, NetElementContainer, False, ActivityTypeSet1)
    FlowBundle.ExecuteCurrentConditions()

def createO02(dfPathLegs):
    # Split the data out by movement type for station summaries before Tube data is overwritten
    dfOB = dfPathLegs.loc[(dfPathLegs.MovementType == 'FirstRailLeg')].copy()
    dfDA = dfPathLegs.loc[(dfPathLegs.MovementType == 'LastRailLeg') | ((dfPathLegs.MovementType == 'FirstRailLeg') & (dfPathLegs.PATHINDEX != dfPathLegs.PATHINDEX.shift(-1)))].copy()
    dfRT = dfPathLegs.loc[dfPathLegs.MovementType == 'RailTransfer'].copy()
    dfTT= dfPathLegs.loc[dfPathLegs.MovementType == 'TubeTransfer'].copy()

    # Next work on station summaries

    #    Start with origin boards
    dfOB['ToPlatform'] = dfOB['FromPlatform']
    dfOB['FromPlatform'] = 'Entry'
    dfOB['ToCRS'] = dfOB.FromCRS
    dfOB = dfOB[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]

    #    Then destination alights. This needs the times recalculating as passengers will appear in the station at the end of the leg
    dfDA['FromPlatform'] = dfDA['ToPlatform']
    dfDA['ToPlatform'] = 'Exit'
    dfDA['FromCRS'] = dfDA['ToCRS']
    dfDA['DEPTIME'] = dfDA.DEPTIME + pd.to_timedelta(dfDA.TIME, "min")
    dfDA['Hour'] = dfDA.DEPTIME.dt.hour
    dfDA = dfDA[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]

    #   Then rail transfers
    dfRT = dfRT[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]

    #   And finally the tube transfers, first the start station movement 
    dfTTFrom = dfTT.copy()
    dfTTFrom['ToPlatform'] = 'Tube'
    dfTTFrom = dfTTFrom[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]

    #   And then the end station movement, which again needs the times recalculating
    dfTTTo = dfTT.copy()
    dfTTTo['FromPlatform'] = 'Tube'
    dfTTTo['DEPTIME'] = dfTTTo.DEPTIME + pd.to_timedelta(dfTTTo.TIME, "min")
    dfTTTo['Hour'] = dfTTTo.DEPTIME.dt.hour
    dfTTTo = dfTTTo[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]

    dfStations = pd.concat([dfOB, dfDA, dfRT, dfTTFrom, dfTTTo], ignore_index=True)

    del dfOB
    del dfDA
    del dfRT
    del dfTTFrom
    del dfTTTo
    del dfTT

    dfStations = dfStations.groupby(['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour'], as_index=False).ODTRIPS.sum()

    dfStations.to_parquet('O02_StationMovements.parquet', index=False, compression=parquetCompression)

    del dfStations


def create_O03(dfPathLegs):
    dfPathLegs.FromCRS = dfPathLegs.FromCRS.astype(str)

    dfDemand = dfPathLegs.groupby(['OrigCRS', 'DestCRS', 'PATHINDEX'], as_index=False).agg(Demand=('ODTRIPS', np.mean),FromCRS=('FromCRS',",".join), ATOC=('ATOC',','.join), StartHour=('Hour',np.min), EndHour=('Hour',np.max), Time=('TIME', np.sum), WaitTime=('WAITTIME', np.sum))

    del dfPathLegs

    dfDemand.rename({'FromCRS':'CRS_Chain', 'ATOC':'ATOC_Chain'}, axis=1, inplace=True)
    dfDemand.CRS_Chain = dfDemand.CRS_Chain + "," + dfDemand.DestCRS

    dfHighLevel = dfDemand.groupby(['OrigCRS', 'DestCRS', 'StartHour', 'EndHour', 'CRS_Chain', 'ATOC_Chain'], as_index=False).agg(Demand=('Demand', np.sum), InVehicleTime=('Time', np.mean), WaitTime=('WaitTime', np.mean))
    dfHighLevel.to_parquet('O03_ODHourlyRoutes.parquet', index=False, compression=parquetCompression)


def create_O04():

    VJI_list = Visum.Workbench.Lists.CreateVehJourneyItemList
    for col in ["VEHJOURNEYNO", "INDEX", r"VEHJOURNEY\TRAINUID", r"VEHJOURNEY\ATOC", r"VEHJOURNEY\TRAINSERVICECODE", r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOP\CRS", r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOP\NAME", r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\NAME", r"EXTARRIVAL", r"EXTDEPARTURE", r"TIMEPROFILEITEM\ALIGHT", r"TIMEPROFILEITEM\BOARD", r"VEHJOURNEY\FROMSTOPPOINT\STOPAREA\STOP\CRS", r"VEHJOURNEY\TOSTOPPOINT\STOPAREA\STOP\CRS", "PASSBOARD(AP)", "PASSALIGHT(AP)", "PASSTHROUGH(AP)"]:
        VJI_list.AddColumn(col)
    
    VJs = [int(x[1]) for x in Visum.Net.VehicleJourneys.GetMultiAttValues("NO", False)]

    VJI_list.SetObjects(False, VJs)

    dfVJIs = pd.DataFrame(VJI_list.SaveToArray(), columns=["VehicleJourneyNo", "Index", "TrainUID", "ATOC", "TrainServiceCode", "CRS", "Stop", "Platform", "Arrival", "Departure", "AlightAllowed", "BoardAllowed", "OriginCRS", "DestinationCRS", "Board", "Alight", "Through"])
    dfVJIs.to_parquet("O04_StopsAndPasses.parquet", index=False, compression=parquetCompression)


    
def create_O05(tempPath):

    Visum.Filters.VolumeAttributeValueFilter().FilterByActiveODPairsAndPuTPaths = False
    cond = Visum.Filters.ODPairFilter().AddCondition("OP_NONE", False, "TOTAL_DEMAND", 3, 0)
    Visum.Filters.ODPairFilter().UseFilter = True

    OD_list = Visum.Workbench.Lists.CreateODPairList
    OD_list.SetObjects(True)

    for col in ["FROMZONE\CODE", "TOZONE\CODE", "TOTAL_DEMAND", "MATVALUE(8)", "MATVALUE(9)", "MATVALUE(10)", "MATVALUE(17)", "MATVALUE(18)", "MATVALUE(19)", "MATVALUE(25)", "MATVALUE(50)", "MATVALUE(33)", "MATVALUE(34)", "MATVALUE(35)", "MATVALUE(42)", "MATVALUE(43)", "MATVALUE(44)", "MATVALUE(58)", "MATVALUE(59)", "MATVALUE(60)", "MATVALUE(67)", "MATVALUE(68)", "MATVALUE(69)"]:
        OD_list.AddColumn(col)

    timecode = datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')

    OD_list.SaveToSQLiteDatabase(f"{tempPath}\\OD_Pairs_{timecode}.sqlite3", "OD_Pairs")

    SQL_Query = 'SELECT "FROMZONE\CODE", "TOZONE\CODE", "TOTAL_DEMAND", "MATVALUE(8)", "MATVALUE(9)", "MATVALUE(10)", "MATVALUE(17)", "MATVALUE(18)", "MATVALUE(19)", "MATVALUE(25)", "MATVALUE(50)", "MATVALUE(33)", "MATVALUE(34)", "MATVALUE(35)", "MATVALUE(42)", "MATVALUE(43)", "MATVALUE(44)", "MATVALUE(58)", "MATVALUE(59)", "MATVALUE(60)", "MATVALUE(67)", "MATVALUE(68)", "MATVALUE(69)" FROM OD_Pairs'

    con = sqlite3.connect(f"{tempPath}\\OD_Pairs_{timecode}.sqlite3") 
    dfODs = pd.read_sql_query(SQL_Query, con)# , chunksize=10000

    con = None

    dfODs.rename({"FROMZONE\CODE":'FromCRS', 'TOZONE\CODE':'ToCRS', 'MATVALUE(8)':'Demand_7-8', 'MATVALUE(9)':'Demand_8-9', 'MATVALUE(10)':'Demand_9-10', 'MATVALUE(17)':'Demand_16-17', 'MATVALUE(18)':'Demand_17-18', 'MATVALUE(19)':'Demand_18-19', 'MATVALUE(25)':'JRT_24hr', 'MATVALUE(50)':'PJT_24hr', 'MATVALUE(33)':'JRT_7-8', 'MATVALUE(34)':'JRT_8-9', 'MATVALUE(35)':'JRT_9-10', 'MATVALUE(42)':'JRT_16-17', 'MATVALUE(43)':'JRT_17-18', 'MATVALUE(44)':'JRT_18-19', 'MATVALUE(58)':'PJT_7-8', 'MATVALUE(59)':'PJT_8-9', 'MATVALUE(60)':'PJT_9-10', 'MATVALUE(67)':'PJT_16-17', 'MATVALUE(68)':'PJT_17-18', 'MATVALUE(69)':'PJT_18-19'})

    dfODs.to_parquet("O05_DemandAndSkims.parquet", index=False, compression=parquetCompression)

    Path.unlink(Path(f"{tempPath}\\OD_Pairs_{timecode}.sqlite3"))    



def main():

    #* This script can be run either inside Visum or externally. If running externally, the path to the version file should be provided in the LoadVersion step of the 'if' block below
    #* Running externally will reduce memory consumption as the version file can be closed before the most memory-intensive steps

    quitVisum = False

    if 'Visum' not in globals():
        import win32com.client as com
        import wx
        
        global app
        app = wx.App()
        
        global Visum
        Visum = com.Dispatch("Visum.Visum.230")
        Visum.LoadVersion(r"C:\Users\david.aspital\PTV Group\Team Network Model T2BAU - General\07 Model Files\19 M16 May23\M16_31May23NC_Assigned.ver")

        quitVisum = True


    tempPath = f"{tempfile.gettempdir()}\\SRAM"
    Path(tempPath).mkdir(parents=True, exist_ok=True)


    flowBundle = False
    #* Change this for station of interest if flowBundle = True
    CRS = 'SRA'

    if flowBundle:
        runFlowBundle(CRS)

    create_O04()
    create_O05(tempPath)
        
    dfStopPoints = getStopPoints(stopPointCols)
    dfPathLegs, timecode = getPathLegs(pathLegCols, tempPath, flowBundle, quitVisum)
    
    dfPathLegs = dfPathLegs.merge(dfStopPoints, left_on='FROMSTOPPOINTNO', right_on='NO', how='left')
    dfPathLegs.drop(['NO', 'FROMSTOPPOINTNO'], axis=1, inplace=True)
    dfPathLegs.rename({'CODE':'FromCode', 'NAME':'FromPlatform', 'CRS':'FromCRS'}, axis=1, inplace=True)

    dfPathLegs = dfPathLegs.merge(dfStopPoints, left_on='TOSTOPPOINTNO', right_on='NO', how='left')
    dfPathLegs.drop(['NO', 'TOSTOPPOINTNO'], axis=1, inplace=True)
    dfPathLegs.rename({'CODE':'ToCode', 'NAME':'ToPlatform', 'CRS':'ToCRS'}, axis=1, inplace=True)

    del dfStopPoints

    dfPathLegs.ToCRS.fillna(dfPathLegs.DestCRS, inplace=True)
    dfPathLegs.FromCRS.fillna(dfPathLegs.OrigCRS, inplace=True)
    dfPathLegs.ToCode.fillna(dfPathLegs.DestCRS, inplace=True)
    dfPathLegs.ToPlatform.fillna("Tube", inplace=True)

    dfPathLegs['Hour'] = dfPathLegs.DEPTIME.dt.hour

    createO02(dfPathLegs)

    for col in ['FromCode', 'ToCode']:
        dfPathLegs[col] = np.where(dfPathLegs.MovementType=='TubeTransfer',
                                dfPathLegs[col].str.split("_").str[0],
                                dfPathLegs[col])
    
    for col in ['FromPlatform', 'ToPlatform']:
        dfPathLegs[col] = np.where(dfPathLegs.MovementType=='TubeTransfer',
                                'Tube',
                                dfPathLegs[col])

    dfPathLegs = dfPathLegs[['OrigCRS', 
                            'DestCRS', 
                            'PATHINDEX', 
                            'MovementType', 
                            'DEPTIME', 
                            'Hour', 
                            'TIME', 
                            'WAITTIME', 
                            'ODTRIPS', 
                            'FromCode', 
                            'FromPlatform', 
                            'FromCRS', 
                            'ToCode', 
                            'ToPlatform', 
                            'ToCRS', 
                            'TrainUID', 
                            'TrainServiceCode', 
                            'ATOC'
                            ]]

    dfPathLegs.to_parquet("O01_PathLegs.parquet", index=False, compression=parquetCompression)

    create_O03(dfPathLegs)
    
    Path.unlink(Path(f"{tempPath}\\PuTPathLegs_{timecode}.sqlite3"))
    print("Done")


if __name__ == '__main__':
    main()
import sys
import os

sys.path.append(os.path.join(os.path.dirname(__file__), "src"))
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__))))


import pandas as pd
import numpy as np
import datetime
import tempfile
from pathlib import Path
import sqlite3

from zipfile import ZipFile
import zipfile

try:
    import get_inputs as gi
    partOfFullRun = True
except:
    partOfFullRun = False


#* Set parquet compression type here
#* From: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_parquet.html:
#* Name of the compression to use. Use None for no compression. Supported options: ‘snappy’, ‘gzip’, ‘brotli’, ‘lz4’, ‘zstd’.
parquetCompression = 'snappy'

pathLegCols = {"PATH\ORIGCONNECTOR\ZONE\CODE":str,
               "PATH\DESTCONNECTOR\ZONE\CODE":str,
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

    SQL_Query = 'SELECT PATHINDEX, PATHLEGINDEX, "PATH\ORIGCONNECTOR\ZONE\CODE", "PATH\DESTCONNECTOR\ZONE\CODE", "PATH\COUNT:PUTPATHLEGSWITHWALK", ODTRIPS, FROMSTOPPOINTNO, TOSTOPPOINTNO, TIMEPROFILEKEYSTRING, TIME, WAITTIME, "STARTVEHJOURNEYITEM\VEHJOURNEY\ATOC", "STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINUID", "STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINSERVICECODE", DEPTIME FROM PathLegs WHERE TIMEPROFILEKEYSTRING NOT IN ("Origin connector", "Destination connector") AND PATHLEGINDEX != 0'

    con = sqlite3.connect(f"{tempPath}\\PuTPathLegs_{timecode}.sqlite3") 
    dfPathLegs = pd.read_sql_query(SQL_Query, con, dtype=cols)# , chunksize=10000

    dfPathLegs.rename({"PATH\ORIGCONNECTOR\ZONE\CODE":'OrigCRS', 'PATH\DESTCONNECTOR\ZONE\CODE':'DestCRS', 'STARTVEHJOURNEYITEM\VEHJOURNEY\ATOC':'ATOC', 'STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINUID':'TrainUID', 'STARTVEHJOURNEYITEM\VEHJOURNEY\TRAINSERVICECODE':'TrainServiceCode', "PATH\COUNT:PUTPATHLEGSWITHWALK":"NumLegs" }, axis=1, inplace=True, errors='ignore')

    for col in ['FROMSTOPPOINTNO', 'TOSTOPPOINTNO']:
        dfPathLegs[col].fillna(-1, inplace=True)
        dfPathLegs[col] = dfPathLegs[col].astype(int)


    dfPathLegs['TIME'] = np.where((dfPathLegs.TIMEPROFILEKEYSTRING == 'Transfer') & ( dfPathLegs.TIMEPROFILEKEYSTRING.shift(-1) == 'PuT Aux PuTAux'),
                                  dfPathLegs.TIME+dfPathLegs.TIME.shift(-1),
                                  dfPathLegs.TIME)
    
    dfPathLegs['TIME'] = np.where((dfPathLegs.TIMEPROFILEKEYSTRING == 'Transfer') & ( dfPathLegs.TIMEPROFILEKEYSTRING.shift(1) == 'PuT Aux PuTAux'),
                                  dfPathLegs.TIME+dfPathLegs.TIME.shift(1),
                                  dfPathLegs.TIME)

    dfPathLegs = dfPathLegs.drop(dfPathLegs[dfPathLegs.TIMEPROFILEKEYSTRING=='PuT Aux PuTAux'].index)

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
                                          np.where(dfPathLegs.TIMEPROFILEKEYSTRING == "Transfer",
                                                   np.where((dfPathLegs.PATHLEGINDEX==3)&(dfPathLegs.FROMSTOPPOINTNO==-1),
                                                            'OriginTubeTransfer',
                                                            np.where((dfPathLegs.PATHLEGINDEX==dfPathLegs.NumLegs-2)&(dfPathLegs.TOSTOPPOINTNO==-1),
                                                                     'DestinationTubeTransfer',
                                                                     np.where(dfPathLegs.TOSTOPPOINTNO==-1,
                                                                              'ToTubeTransfer',
                                                                              np.where(dfPathLegs.FROMSTOPPOINTNO==-1,
                                                                                       'FromTubeTransfer',
                                                                                       'RailTransfer'
                                                                                       )
                                                                              ),
                                                                    ),
                                                            ),
                                                   np.where(dfPathLegs.PATHLEGINDEX==dfPathLegs.NumLegs-1,
                                                            'LastRailLeg',
                                                            'IntermediateRailLeg'
                                                            )
                                                   )
                                          )
                                         
    
    dfPathLegs.drop(['NumLegs', 'TIMEPROFILEKEYSTRING'], axis=1, inplace=True)
    
    tubeMovements = ['FromTubeTransfer', 'ToTubeTransfer', 'OriginTubeTransfer', 'DestinationTubeTransfer']

    dfPathLegs['ATOC'] = np.where(dfPathLegs.MovementType.isin(tubeMovements),
                                  'Tube',
                                  dfPathLegs.ATOC)

    dfPathLegs.drop(['PATHLEGINDEX'], axis=1, inplace=True)

    con.close()
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

def createO02(dfPathLegs, runID):
    # Split the data out by movement type for station summaries before Tube data is overwritten
    dfOB = dfPathLegs.loc[(dfPathLegs.MovementType == 'FirstRailLeg')].copy()
    dfDA = dfPathLegs.loc[(dfPathLegs.MovementType == 'LastRailLeg') | ((dfPathLegs.MovementType == 'FirstRailLeg') & (dfPathLegs.PATHINDEX != dfPathLegs.PATHINDEX.shift(-1)))].copy()
    dfRT = dfPathLegs.loc[dfPathLegs.MovementType == 'RailTransfer'].copy()
    dfFT = dfPathLegs.loc[dfPathLegs.MovementType == 'FromTubeTransfer'].copy()
    dfTT = dfPathLegs.loc[dfPathLegs.MovementType == 'ToTubeTransfer'].copy()
    dfOT = dfPathLegs.loc[dfPathLegs.MovementType == 'OriginTubeTransfer'].copy()
    dfDT = dfPathLegs.loc[dfPathLegs.MovementType == 'DestinationTubeTransfer'].copy()


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

    #   And finally the tube transfers, starting with the to tube movements
    dfTT['ToCRS'] = dfTT['FromCRS']
    dfTT = dfTT[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]
    
    # Origin transfers, to tube movement. So FromPlatform is station entry, ToPlatform is Tube, set ToCRS to FromCRS
    dfOT_To = dfOT.copy()
    dfOT_To['FromPlatform'] = 'Entry'
    dfOT_To['ToCRS'] = dfOT_To['FromCRS']
    dfOT_To['ToPlatform'] = 'Tube'
    dfOT_To = dfOT_To[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]
    
    # Destination transfers, to tube movement. So FromPlatform is the same, set ToPlatform to Tube, set ToCRS to FromCRS
    dfDT_To = dfDT.copy()
    dfDT_To['ToCRS'] = dfDT_To['FromCRS']
    dfDT_To['ToPlatform'] = 'Tube'
    dfDT_To = dfDT_To[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]

    #   And then the from tube movements, which again needs the times recalculating
    dfFT['FromCRS'] = dfFT['ToCRS']
    dfFT['DEPTIME'] = dfFT.DEPTIME + pd.to_timedelta(dfFT.TIME, "min")
    dfFT['Hour'] = dfFT.DEPTIME.dt.hour
    dfFT = dfFT[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]

    # Origin transfers, from tube movement. So FromPlatform is Tube, ToPlatform is the same, set FromCRS to ToCRS. Recalculate DepartureTime and Hour
    dfOT_From = dfOT.copy()
    del dfOT
    dfOT_From['FromCRS'] = dfOT_From['ToCRS'] 
    dfOT_From['DEPTIME'] = dfOT_From.DEPTIME + pd.to_timedelta(dfOT_From.TIME, "min")
    dfOT_From['Hour'] = dfOT_From.DEPTIME.dt.hour
    dfOT_From = dfOT_From[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]
    
    # Destination transfers, from tube movement. FromPlatform is tube, ToPlatform is Exit. Set FromCRS to ToCRS. Recalculate DepartureTime and Hour
    dfDT_From = dfDT.copy()
    del dfDT
    dfDT_From['FromCRS'] = dfDT_From['ToCRS'] 
    dfDT_From['FromPlatform'] = 'Tube'
    dfDT_From['ToPlatform'] = 'Exit'
    dfDT_From['DEPTIME'] = dfDT_From.DEPTIME + pd.to_timedelta(dfDT_From.TIME, "min")
    dfDT_From['Hour'] = dfDT_From.DEPTIME.dt.hour
    dfDT_From = dfDT_From[['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour', 'ODTRIPS']]


    dfStations = pd.concat([dfOB, dfDA, dfRT, dfFT, dfTT, dfOT_To, dfOT_From, dfDT_From, dfDT_To], ignore_index=True)

    del dfOB
    del dfDA
    del dfRT
    del dfFT
    del dfTT
    del dfOT_To
    del dfOT_From
    del dfDT_From
    del dfDT_To

    dfStations = dfStations.groupby(['FromCRS', 'FromPlatform', 'ToCRS', 'ToPlatform', 'Hour'], as_index=False).ODTRIPS.sum()

    dfStations['RunID'] = runID

    dfStations.to_parquet(f'{runID}_O02_StationMovements.parquet', index=False, compression=parquetCompression)
    dfStations.to_csv(f'{runID}_O02_StationMovements.csv', index=False)

    del dfStations


def create_O03(dfPathLegs, runID):
    dfPathLegs.FromCRS = dfPathLegs.FromCRS.astype(str)

    dfPathLegs.drop(dfPathLegs[dfPathLegs.MovementType=='FromTubeTransfer'].index, inplace=True)

    dfDemand = dfPathLegs.groupby(['OrigCRS', 'DestCRS', 'PATHINDEX'], as_index=False).agg(Demand=('ODTRIPS', np.mean),FromCRS=('FromCRS',",".join), ATOC=('ATOC',','.join), StartHour=('Hour',np.min), EndHour=('Hour',np.max), Time=('TIME', np.sum), WaitTime=('WAITTIME', np.sum))

    del dfPathLegs

    dfDemand.rename({'FromCRS':'CRS_Chain', 'ATOC':'ATOC_Chain'}, axis=1, inplace=True)
    dfDemand.CRS_Chain = dfDemand.CRS_Chain + "," + dfDemand.DestCRS

    dfHighLevel = dfDemand.groupby(['OrigCRS', 'DestCRS', 'StartHour', 'EndHour', 'CRS_Chain', 'ATOC_Chain'], as_index=False).agg(Demand=('Demand', np.sum), InVehicleTime=('Time', np.mean), WaitTime=('WaitTime', np.mean))
    dfHighLevel['RunID'] = runID
    dfHighLevel.to_parquet(f'{runID}_O03_ODHourlyRoutes.parquet', index=False, compression=parquetCompression)
    dfHighLevel.to_csv(f'{runID}_O03_ODHourlyRoutes.csv', index=False)


def create_O04(runID):

    VJI_list = Visum.Workbench.Lists.CreateVehJourneyItemList
    for col in ["VEHJOURNEYNO", "INDEX", r"VEHJOURNEY\TRAINUID", r"VEHJOURNEY\ATOC", r"VEHJOURNEY\MeanVolTrip(AP)", r"VEHJOURNEY\TRAINSERVICECODE", r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOP\CODE" ,r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOP\CRS", r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\STOPAREA\STOP\NAME", r"TIMEPROFILEITEM\LINEROUTEITEM\STOPPOINT\NAME", r"EXTARRIVAL", r"EXTDEPARTURE", r"TIMEPROFILEITEM\ALIGHT", r"TIMEPROFILEITEM\BOARD", r"VEHJOURNEY\FROMSTOPPOINT\STOPAREA\STOP\CRS", r"VEHJOURNEY\TOSTOPPOINT\STOPAREA\STOP\CRS", "PASSBOARD(AP)", "PASSALIGHT(AP)", "PASSTHROUGH(AP)"]:
        VJI_list.AddColumn(col)
    
    VJs = [int(x[1]) for x in Visum.Net.VehicleJourneys.GetMultiAttValues("NO", False)]

    VJI_list.SetObjects(False, VJs)

    dfVJIs = pd.DataFrame(VJI_list.SaveToArray(), columns=["VehicleJourneyNo", "Index", "TrainUID", "ATOC", "MeanVolTrip(AP)", "TrainServiceCode", "CODE", "CRS", "Stop", "Platform", "Arrival", "Departure", "AlightAllowed", "BoardAllowed", "OriginCRS", "DestinationCRS", "Board", "Alight", "Through"])
    dfVJIs['RunID'] = runID
    dfVJIs.to_parquet(f"{runID}_O04_StopsAndPasses.parquet", index=False, compression=parquetCompression)
    dfVJIs.to_csv(f"{runID}_O04_StopsAndPasses.csv", index=False)


    
def create_O05(tempPath, runID):

    Visum.Filters.VolumeAttributeValueFilter().FilterByActiveODPairsAndPuTPaths = False
    Visum.Filters.ODPairFilter().Init()
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
    con.close()

    con = None

    dfODs.rename({"FROMZONE\CODE":'FromCRS', 'TOZONE\CODE':'ToCRS', 'MATVALUE(8)':'Demand_7-8', 'MATVALUE(9)':'Demand_8-9', 'MATVALUE(10)':'Demand_9-10', 'MATVALUE(17)':'Demand_16-17', 'MATVALUE(18)':'Demand_17-18', 'MATVALUE(19)':'Demand_18-19', 'MATVALUE(25)':'JRT_24hr', 'MATVALUE(50)':'PJT_24hr', 'MATVALUE(33)':'JRT_7-8', 'MATVALUE(34)':'JRT_8-9', 'MATVALUE(35)':'JRT_9-10', 'MATVALUE(42)':'JRT_16-17', 'MATVALUE(43)':'JRT_17-18', 'MATVALUE(44)':'JRT_18-19', 'MATVALUE(58)':'PJT_7-8', 'MATVALUE(59)':'PJT_8-9', 'MATVALUE(60)':'PJT_9-10', 'MATVALUE(67)':'PJT_16-17', 'MATVALUE(68)':'PJT_17-18', 'MATVALUE(69)':'PJT_18-19'}, axis=1, inplace=True)

    dfODs['RunID'] = runID
    dfODs.to_parquet(f"{runID}_O05_DemandAndSkims.parquet", index=False, compression=parquetCompression)
    dfODs.to_csv(f"{runID}_O05_DemandAndSkims.csv", index=False)

    Path.unlink(Path(f"{tempPath}\\OD_Pairs_{timecode}.sqlite3"))    

def create_O06(tempPath, runID):
    OD_list = Visum.Workbench.Lists.CreateODPairList
    OD_list.SetObjects(True)

    for col in ["FROMZONE\CODE", "TOZONE\CODE", "MATVALUE(26)", "MATVALUE(27)", "MATVALUE(28)", "MATVALUE(29)", "MATVALUE(30)", "MATVALUE(31)", "MATVALUE(32)", "MATVALUE(33)", "MATVALUE(34)", "MATVALUE(35)", "MATVALUE(36)", "MATVALUE(37)", "MATVALUE(38)", "MATVALUE(39)", "MATVALUE(40)", "MATVALUE(41)", "MATVALUE(42)", "MATVALUE(43)", "MATVALUE(44)", "MATVALUE(45)", "MATVALUE(46)", "MATVALUE(47)", "MATVALUE(48)", "MATVALUE(49)"]:
        OD_list.AddColumn(col)

    timecode = datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')

    OD_list.SaveToSQLiteDatabase(f"{tempPath}\\OD_Pairs_{timecode}.sqlite3", "OD_Pairs")

    SQL_Query = 'SELECT "FROMZONE\CODE", "TOZONE\CODE", "MATVALUE(26)", "MATVALUE(27)", "MATVALUE(28)", "MATVALUE(29)", "MATVALUE(30)", "MATVALUE(31)", "MATVALUE(32)", "MATVALUE(33)", "MATVALUE(34)", "MATVALUE(35)", "MATVALUE(36)", "MATVALUE(37)", "MATVALUE(38)", "MATVALUE(39)", "MATVALUE(40)", "MATVALUE(41)", "MATVALUE(42)", "MATVALUE(43)", "MATVALUE(44)", "MATVALUE(45)", "MATVALUE(46)", "MATVALUE(47)", "MATVALUE(48)", "MATVALUE(49)" FROM OD_Pairs'

    con = sqlite3.connect(f"{tempPath}\\OD_Pairs_{timecode}.sqlite3") 
    dfODs = pd.read_sql_query(SQL_Query, con)# , chunksize=10000
    con.close()
    con = None

    dfODs.rename({"FROMZONE\CODE":'FromCRS', 'TOZONE\CODE':'ToCRS', "MATVALUE(26)":'JRT_0-1', "MATVALUE(27)":'JRT_1-2', "MATVALUE(28)":'JRT_2-3', "MATVALUE(29)":'JRT_3-4', "MATVALUE(30)":'JRT_4-5', "MATVALUE(31)":'JRT_5-6', "MATVALUE(32)":'JRT_6-7', "MATVALUE(33)":'JRT_7-8', "MATVALUE(34)":'JRT_8-9', "MATVALUE(35)":'JRT_9-10', "MATVALUE(36)":'JRT_10-11', "MATVALUE(37)":'JRT_11-12', "MATVALUE(38)":'JRT_12-13', "MATVALUE(39)":'JRT_13-14', "MATVALUE(40)":'JRT_14-15', "MATVALUE(41)":'JRT_15-16', "MATVALUE(42)":'JRT_16-17', "MATVALUE(43)":'JRT_17-18', "MATVALUE(44)":'JRT_18-19', "MATVALUE(45)":'JRT_19-20', "MATVALUE(46)":'JRT_20-21', "MATVALUE(47)":'JRT_21-22', "MATVALUE(48)":'JRT_22-23', "MATVALUE(49)":'JRT_23-24'}, axis=1, inplace=True)

    dfODs['RunID'] = runID
    dfODs.to_parquet(f"{runID}_O06_JRTSkims.parquet", index=False, compression=parquetCompression)
    dfODs.to_csv(f"{runID}_O06_JRTSkims.csv", index=False)

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
        Visum = com.Dispatch("Visum.Visum.240")
        Visum.LoadVersion(r"C:\Users\david.aspital\PTV Group\Team Network Model T2BAU - General\07 Model Files\32 M42\M42_31May23_Assigned.ver")

        quitVisum = True


    tempPath = f"{tempfile.gettempdir()}\\SRAM"
    Path(tempPath).mkdir(parents=True, exist_ok=True)


    flowBundle = False
    #* Change this for station of interest if flowBundle = True
    CRS = 'LBG'

    if flowBundle:
        runFlowBundle(CRS)


    if partOfFullRun:
        path = os.path.dirname(os.path.dirname(__file__))
        input_path = os.path.join(path, "input\\inputs.csv")
        runID = gi.getRunID(input_path)
    else:
        runID = os.path.split(Visum.IO.CurrentVersionFile)[1][:3]

    create_O04(runID)
    create_O05(tempPath, runID)
    create_O06(tempPath, runID)
        
    dfStopPoints = getStopPoints(stopPointCols)
    dfPathLegs, timecode = getPathLegs(pathLegCols, tempPath, flowBundle, quitVisum)
    
    dfPathLegs = dfPathLegs.merge(dfStopPoints, left_on='FROMSTOPPOINTNO', right_on='NO', how='left')
    dfPathLegs.drop(['NO', 'FROMSTOPPOINTNO'], axis=1, inplace=True)
    dfPathLegs.rename({'CODE':'FromCode', 'NAME':'FromPlatform', 'CRS':'FromCRS'}, axis=1, inplace=True)
    dfPathLegs.FromCode.fillna('Tube', inplace=True)
    dfPathLegs.FromPlatform.fillna('Tube', inplace=True)
    dfPathLegs.FromCRS.fillna(dfPathLegs.OrigCRS, inplace=True)

    dfPathLegs = dfPathLegs.merge(dfStopPoints, left_on='TOSTOPPOINTNO', right_on='NO', how='left')
    dfPathLegs.drop(['NO', 'TOSTOPPOINTNO'], axis=1, inplace=True)
    dfPathLegs.rename({'CODE':'ToCode', 'NAME':'ToPlatform', 'CRS':'ToCRS'}, axis=1, inplace=True)
    dfPathLegs.ToCode.fillna('Tube', inplace=True)
    dfPathLegs.ToPlatform.fillna('Tube', inplace=True)
    dfPathLegs.ToCRS.fillna(dfPathLegs.DestCRS, inplace=True)

    del dfStopPoints

    dfPathLegs['Hour'] = dfPathLegs.DEPTIME.dt.hour

    createO02(dfPathLegs, runID)

    tubeMovements = ['FromTubeTransfer', 'ToTubeTransfer', 'OriginTubeTransfer', 'DestinationTubeTransfer']

    for col in ['FromCode', 'ToCode']:
        dfPathLegs[col] = np.where(dfPathLegs.MovementType.isin(tubeMovements),
                                dfPathLegs[col].str.split("_").str[0],
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
    dfPathLegs['RunID'] = runID
    dfPathLegs.to_parquet(f"{runID}_O01_PathLegs.parquet", index=False, compression=parquetCompression)
    dfPathLegs.to_csv(f"{runID}_O01_PathLegs.csv", index=False)

    create_O03(dfPathLegs, runID)
    
    Path.unlink(Path(f"{tempPath}\\PuTPathLegs_{timecode}.sqlite3"))
    print("Done")

    path = os.path.dirname(__file__)
    files = [f"{runID}_O01_PathLegs.csv", f'{runID}_O02_StationMovements.csv', f'{runID}_O03_ODHourlyRoutes.csv', f"{runID}_O04_StopsAndPasses.csv", f"{runID}_O05_DemandAndSkims.csv", f"{runID}_O06_JRTSkims.csv"]

    with ZipFile(f'{runID}_Results.zip', 'w',  zipfile.ZIP_DEFLATED) as zipObj:
    # Iterate over all the files in directory

            for file in files:
                zipObj.write(os.path.join(path, file), compresslevel=9)
                Path.unlink(os.path.join(path, file))



if __name__ == '__main__':
    main()
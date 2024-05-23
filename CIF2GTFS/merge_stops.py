import win32com.client
import os
import pandas as pd
import numpy as np
import datetime
import tempfile
import traceback

import sys
sys.path.append(os.path.dirname(__file__))

import get_inputs as gi


att_header = '''$VISION
* {}
* {}
* 
* Table: Version block
* 
$VERSION:VERSNR	FILETYPE	LANGUAGE	UNIT
13	Att	ENG	KM

* 
* Table: {}s
* 

'''

def export_att_file(df, folder, filename, visum_object="TRANSFERSANDWALKTIMESWITHINSTOP"):
    # Export pandas df as .att file that can be ready into Visum

    # Check for extension 
    if filename[-4:] == ".att":
        pass
    else:
        filename = filename+".att"
        
    # Write .att header
    with open (folder+r"\\"+filename,'w') as o:
        o.write(att_header.format(os.getlogin(), datetime.datetime.now(), visum_object.title()))
        o.close()
        # Write dataframe
        df.to_csv(folder+r"\\"+filename, mode='a', sep="\t", index=False)


def df2visum(Visum, df, temp_path, filename):
    '''
    Import data from a DataFrame into Visum. Saves as and att file to a temp location and reads back in
    
    Parameters
    ----------
    df : DataFrame
        DataFrame to be imported into Visum
    temp_path : str, Path object
        Temp location to write att file to
    filename : str
        Filename for the att file
    '''
    export_att_file(df, temp_path, f"{filename}.att")
    Visum.IO.LoadAttributeFile(f"{temp_path}\\{filename}.att")


def merge_stops(merge_path, ver_path):

    Visum = win32com.client.gencache.EnsureDispatch('Visum.Visum.230')
    Visum.SetPath(57, os.path.join(path,f"cached_data"))
    Visum.SetLogFileName(f"Log_MergeStops_{datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')}.txt")
    C = win32com.client.constants
    Visum.LoadVersion(ver_path)

    try:
        MergedStops = pd.read_csv(merge_path)
        UniqueStops = MergedStops.drop(MergedStops[MergedStops.NewCoordinates == 0].index)
        OldStopCode = [ stop for stop in UniqueStops["StopCode"]]
        UniqueStopsdict = UniqueStops.set_index('StopCode').T.to_dict('list')
        AllStopsdict = dict(Visum.Net.Stops.GetMultipleAttributes(['CODE', 'NO']))
        
        OldStopNo  = [AllStopsdict[stop] for stop in OldStopCode]
        oldStopKey = [ Visum.Net.Stops.ItemByKey(stop) for stop in OldStopNo ]
        oldStopCordX = [ item.AttValue("XCOORD") for item in oldStopKey ]
        oldStopCordY = [ item.AttValue("YCOORD") for item in oldStopKey ]
        oldStopCoords = zip(oldStopCordX, oldStopCordY)
        oldStopNo = [ item.AttValue("No") for item in oldStopKey ]
        oldStopCord = dict(zip(oldStopCordX, oldStopCordY))

        for i, (key, value) in enumerate(oldStopCord.items()):
            Name = UniqueStopsdict[OldStopCode[i]][0] 
            Code = UniqueStopsdict[OldStopCode[i]][1]
            CRS = UniqueStopsdict[OldStopCode[i]][2]
            FilStops = MergedStops.loc[MergedStops["NewCode"] == Code]
            oldStopCode = list(FilStops["StopCode"]) # list of old stop code 
            oldStopNum =  [AllStopsdict[stop] for stop in oldStopCode] # get a list of old stop numbers 
            oldStopKey = [ Visum.Net.Stops.ItemByKey(stop) for stop in oldStopNum ]
            oldStopArea = [stop.StopAreas.GetAll for stop in oldStopKey] # list of all stops area with the stop numbers
            oldStopAreaNo = [StopArea.AttValue("NO") for item in oldStopArea for StopArea in item]

            if Code != "":
                NewStop = Visum.Net.AddStop(-1, key, value)


                NewStop.SetAttValue("CODE", Code)
                NewStop.SetAttValue("NAME", Name)
                NewStop.SetAttValue("CRS", CRS)
                NewStopNo = NewStop.AttValue("NO")

                AEMoved = False

                for i in oldStopAreaNo:# loop through all stops areas 
                    Newstoparea = Visum.Net.StopAreas.ItemByKey(i)

                    if Newstoparea.AttValue("Name") == "AccessEgress":
                        if not AEMoved:
                            Newstoparea.SetAttValue("StopNo", NewStopNo) # change stop number to the new stop no
                            PlatUnknownSAs = Visum.Net.StopAreas.GetFilteredSet(f'[Name]="Platform Unknown" & [StopNo]={NewStopNo} & [CODE]=[STOP\CODE]')
                            if PlatUnknownSAs.Count > 1:
                                raise Exception(f"More than 1 Platform Unknown for Stop {NewStopNo}")
                            if PlatUnknownSAs.Count == 1:
                                NewNodeNo = PlatUnknownSAs.GetMultipleAttributes(['NodeNo'])[0][0]
                                Newstoparea.SetAttValue("NodeNo", NewNodeNo)
                            AEMoved = True
                        else:
                            Visum.Net.RemoveStopArea(Visum.Net.StopAreas.ItemByKey(i))
                    else:
                        Newstoparea.SetAttValue("StopNo", NewStopNo)

                for i in oldStopNum: # delete the old stops  
                    Visum.Net.RemoveStop(Visum.Net.Stops.ItemByKey(i))

        #read transfer walk time table from Visum
        walkTimeList = Visum.Workbench.Lists.CreateStopTransferWalkTimeList
        for col in ['TIME(W)', 'FROMSTOPAREANO', 'TOSTOPAREANO', r'FROMSTOPAREA\NO', r'FROMSTOPAREA\NAME', r'TOSTOPAREA\NAME']:
            walkTimeList.AddColumn(col)
        walkTimeList.SaveToAttributeFile(os.path.join(tempfile.gettempdir(), "WalkTime.att"), 9)
        walk_time_df = pd.read_csv(os.path.join(tempfile.gettempdir(), "WalkTime.att"), skiprows=12, sep='\t')

        walk_time_df["TIME(W)"] = "1440min"

        filter = ["AccessEgress", "Platform Unknown"]

        walk_time_df["TIME(W)"] = np.where((walk_time_df["FROMSTOPAREA\\NAME"].str.startswith("Platform")) 
                                        & (walk_time_df["TOSTOPAREA\\NAME"].str.startswith("Platform")),"5min",walk_time_df["TIME(W)"])

        walk_time_df["TIME(W)"] = np.where(((~walk_time_df["FROMSTOPAREA\\NAME"].isin(filter)) | (~walk_time_df["TOSTOPAREA\\NAME"].isin(filter))) 
                                        & (walk_time_df["FROMSTOPAREA\\NAME"] == walk_time_df["TOSTOPAREA\\NAME"]),"0min", walk_time_df["TIME(W)"])

        walk_time_df["TIME(W)"] = np.where((walk_time_df["FROMSTOPAREA\\NAME"].str.startswith("Platform")) 
                                                                            & (walk_time_df["TOSTOPAREA\\NAME"] == "AccessEgress"),"0min",walk_time_df["TIME(W)"])

        walk_time_df["TIME(W)"] = np.where((walk_time_df["FROMSTOPAREA\\NAME"] == "AccessEgress") 
                                                                            & (walk_time_df["TOSTOPAREA\\NAME"].str.startswith("Platform")),"0min",walk_time_df["TIME(W)"])

        df2visum(Visum, walk_time_df, tempfile.gettempdir(), "TransferWalkTime.att")

        Visum.SaveVersion(ver_path.replace(".ver", "_MergeStops.ver"))
    except:
        Visum.Log(12288, traceback.format_exc())

def main(merge_path, ver_path):
    merge_stops(merge_path, ver_path)

    

    

if __name__ == '__main__':

    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    mergeStops = gi.readMergeInputs(input_path)

    if mergeStops[0] == "TRUE":
        merge_path = mergeStops[1]
        ver_path = os.path.join(path, "output\\VISUM\\Network+Timetable.ver")

        #merge_path = "CIF2GTFS\\input\\StopsToMerge+CRSOverride.csv"
        main(merge_path, ver_path)
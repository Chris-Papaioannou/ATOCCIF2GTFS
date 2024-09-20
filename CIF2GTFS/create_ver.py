#%%
import os
import numpy as np
import pandas as pd
import win32com.client as com
import json
import datetime
import traceback

import sys
sys.path.append(os.path.dirname(__file__))

import get_inputs as gi

# Format of createVer string: {'Name':'xyz', 'TSysSet':'PO,PX,PE','Date':'31/05/2023'}
# TSysSet can also be blank for all TSys


def main():

    
    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    vers = gi.readVerInputs(input_path)
    if bool(vers[0]):

        #vers=['{"Name":"31May23", "TSysSet":"","Date":"31.05.2023"}']
        for ver in vers[1:]:

            #Launch Visum and load in the final supply network version
            Visum = com.Dispatch('Visum.Visum.240')
            Visum.SetPath(57, os.path.join(path,f"cached_data"))
            Visum.SetLogFileName(f"Log_CreateVers_{datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')}.txt")
            try:
                Visum.IO.LoadVersion(os.path.join(path, 'output\\VISUM\\Network+Timetable_MergeStops.ver'))

                verDict = json.loads(ver)

                verName = verDict['Name']
                verTSys = verDict['TSysSet'].strip()
                verDate = verDict['Date'] #! Format needs to be dd.mm.yyyy

                Visum.Net.CalendarPeriod.SetAttValue('ValidFrom', verDate)
                Visum.Net.CalendarPeriod.SetAttValue('ValidUntil', verDate)

                # Delete all Lines where the VJS are not valid for the date
                invalidLines = Visum.Net.Lines.GetFilteredSet(f'[SUM:LINEROUTES\SUM:VEHJOURNEYS\SUM:VEHJOURNEYSECTIONS\ISVALID({verDate})]=0') # here verDate needs to be dd.mm.yyyy
                invalidLines.RemoveAll()


                TSys_list = verTSys.split(",")
                if TSys_list == ['']:
                    fil_str = '[TSYSCODE]="'+'"|[TSYSCODE]="'.join(TSys_list) +'"'
                    invalidLines = Visum.Net.Lines.GetFilteredSet(fil_str)
                    invalidLines.RemoveAll()     
                
                
                #Then adding the relevant TimeSeries before iterating through the hours, creating matrices, reading in the values, and adding timeSeriesItems
                myTimeSeries = Visum.Net.AddTimeSeries(2,1)

                for i in range(24):
                    myVisumMatrix = Visum.Net.AddMatrixWithFormula(i+1, "1", 2, 3)
                    myVisumMatrix.SetAttValue("CODE", f"Demand {i}-{i+1}")
                    myVisumMatrix.SetAttValue("FromTime", i*60*60)
                    myVisumMatrix.SetAttValue("ToTime", (i+1)*60*60)
                    myTimeSeriesItem = myTimeSeries.AddTimeSeriesItem(3600*i, 3600*(i + 1))
                    myTimeSeriesItem.SetAttValue('Matrix', f"Matrix([NO]={myVisumMatrix.AttValue('No')})")

                myDemandTimeSeries = Visum.Net.DemandTimeSeriesCont.ItemByKey(1)
                myDemandTimeSeries.SetAttValue("TimeSeriesNo",2)

                TI_Set = myTimeSeries.CreateTimeIntervalSetAndConnect()
                
                Visum.Net.CalendarPeriod.SetAttValue('AnalysisTimeIntervalSetNo',1)

                # Create Total_Demand attribute
                Visum.Net.ODPairs.AddUserDefinedAttribute("Total_Demand", "Total_Demand", 'Total_Demand', 2, 2, Formula='[MATVALUE(1)]+[MATVALUE(2)]+[MATVALUE(3)]+[MATVALUE(4)]+[MATVALUE(5)]+[MATVALUE(6)]+[MATVALUE(7)]+[MATVALUE(8)]+[MATVALUE(9)]+[MATVALUE(10)]+[MATVALUE(11)]+[MATVALUE(12)]+[MATVALUE(13)]+[MATVALUE(14)]+[MATVALUE(15)]+[MATVALUE(16)]+[MATVALUE(17)]+[MATVALUE(18)]+[MATVALUE(19)]+[MATVALUE(20)]+[MATVALUE(21)]+[MATVALUE(22)]+[MATVALUE(23)]+[MATVALUE(24)]')

                #Finally save the ver file to assign 
                Visum.IO.SaveVersion(os.path.join(path, f'output\\VISUM\\{verName}.ver'))
            except:
                Visum.Log(12288, traceback.format_exc())


# %%

if __name__ == "__main__":
    myMatrix = main()

# %%
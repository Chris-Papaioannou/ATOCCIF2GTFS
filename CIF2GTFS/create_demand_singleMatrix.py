#%%
import os
import numpy as np
import pandas as pd
import win32com.client as com
import sys
import traceback

sys.path.append(os.path.dirname(__file__))

import get_inputs as gi

import logging

logging.basicConfig(
    filename="ModelBuilder.log",
    encoding="utf-8",
    filemode="a",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M",
    level=logging.INFO # Change to logging.DEBUG for more details
)


def main(demandFilename, CRSUpdate, WeekdayMatrix, MidweekFactors, GroupedStationSplits, TravelcardSplits, TimeProfiles, GlobalFactor):

    logging.info('Creating demand...')

    #Define file path of scripts
    path = os.path.dirname(__file__)

    # 1. Open DailyDemand.att
    # 2. Apply MidWeek Factors
    # 3. Split grouped stations
    # 4. Split travelcard stations 
    # 5. Apply time profiles

    try:

        updateCRS = pd.read_csv(CRSUpdate, keep_default_na=False)
        dictCRS = dict(zip(updateCRS.OldCRS, updateCRS.NewCRS))


        weekdayMat = pd.read_csv(WeekdayMatrix, low_memory = False, names=['Date', 'Orig', 'Dest', 'Hour', 'Demand'], keep_default_na=False, header=0)
        weekdayMat = weekdayMat.loc[weekdayMat.Orig != weekdayMat.Dest].copy()
        weekdayMat.replace({'Orig':dictCRS, 'Dest':dictCRS}, inplace=True)

        hourlyMatrices = weekdayMat.pivot_table('Demand', ['Orig', 'Dest'], 'Hour', fill_value=0).reset_index()

        renameCols = {key: value for key, value in zip(range(24), [f"Matrix({x+1})" for x in range(24)])}

        hourlyMatrices.rename(renameCols, axis=1, inplace=True)
        
        total=0.0
        for i in range(24):
            if f'Matrix({i+1})' in hourlyMatrices.columns.values:
                total += hourlyMatrices[f'Matrix({i+1})'].sum()
        print(total)

        hourlyMatrices.rename({'Orig':'FROMZONE\CODE', 'Dest':'TOZONE\CODE'}, axis=1, inplace=True)
        hourlyMatrices.to_csv(os.path.join(path, 'demand\\HourlyMatrices.csv'),index=False)
        
        Visum = com.Dispatch("Visum.Visum.240")
        Visum.LoadVersion(os.path.join(path, f"output\\VISUM\\{demandFilename}.ver"))
        
        # Get an index of all OD pairs from Visum, and merge our final hourly matrices onto them
        myIndex = pd.DataFrame(Visum.Net.ODPairs.GetMultipleAttributes(['FROMZONE\CODE', 'TOZONE\CODE']), columns = ['FROMZONE\CODE', 'TOZONE\CODE'])
        myExpandedMatrix = myIndex.merge(hourlyMatrices, 'left', ['FROMZONE\CODE', 'TOZONE\CODE'])

        
        #Iterate through the hours, convert dummy matrices to data, set to 0 and read in the values
        for i in range(24):
            myVisumMatrix = Visum.Net.Matrices.ItemByKey(i+1)
            myVisumMatrix.SetAttValue("DATASOURCETYPE", 1)
            myVisumMatrix.SetValuesToResultOfFormula("0")
            if f'Matrix({i+1})' in hourlyMatrices.columns.values:
                myHourlyMatrix = myExpandedMatrix[f'Matrix({i+1})'].reset_index()
                myHourlyMatrix['index'] += 1
                Visum.Net.ODPairs.SetMultiAttValues(f'MatValue({str(i+1)})', myHourlyMatrix.values)
                logging.info(f'Matrix {i+1} imported')


        Visum.IO.SaveVersion(os.path.join(path, f"demand\\{demandFilename}_Demand.ver"))
        
        myExpandedMatrix['OD'] = myExpandedMatrix['FROMZONE\CODE']+"_"+myExpandedMatrix['TOZONE\CODE']
        hourlyMatrices['OD'] = hourlyMatrices['FROMZONE\CODE']+"_"+hourlyMatrices['TOZONE\CODE']

        VisumODs = list(myExpandedMatrix.OD.unique())
        print(len(VisumODs))
        inputODs = list(hourlyMatrices.OD.unique())
        print(len(inputODs))

        missingODs = list(set(inputODs)-set(VisumODs))
        print('Doing loc')
        missingDemand = hourlyMatrices.loc[hourlyMatrices.OD.isin(missingODs)]

        missingDemand.to_csv(os.path.join(path, 'demand\\DroppedDemand.csv'), index=False)
        logging.warning(f'Total dropped demand: {missingDemand.sum()}')
        print('Done')
    except:
        logging.error(traceback.format_exc())

# %%

if __name__ == "__main__":

    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    importDemand = gi.readDemandInputs(input_path)
    logging.info(f'Importing demand: {importDemand[0]}')
    if bool(importDemand[0]):
        demandFilename = importDemand[1]
        CRSUpdate = importDemand[2]
        WeekdayMatrix = importDemand[3]
        MidweekFactors = importDemand[4]
        GroupedStationSplits = importDemand[5]
        TravelcardSplits = importDemand[6]
        TimeProfiles = importDemand[7]
        GlobalFactor = float(importDemand[8])

        logging.info(f'Output version file: {importDemand[1]}')
        logging.info(f'CRS update file: {importDemand[2]}')
        logging.info(f'Daily matrix: {importDemand[3]}')
        logging.info(f'Midweekday factors: {importDemand[4]}')
        logging.info(f'Grouped station splits: {importDemand[5]}')
        logging.info(f'Travelcard splits: {importDemand[6]}')
        logging.info(f'Time profiles: {importDemand[7]}')
        logging.info(f'Global factor: {importDemand[8]}')

        main(demandFilename, CRSUpdate, WeekdayMatrix, MidweekFactors, GroupedStationSplits, TravelcardSplits, TimeProfiles, GlobalFactor)

# %%
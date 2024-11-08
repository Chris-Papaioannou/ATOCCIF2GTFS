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


def applyMidweekFactors(weekdayMat, midweekFactors, globalFactor):

    weekdayMat = weekdayMat.merge(midweekFactors, how='left', left_on=['FromZone', 'ToZone'], right_on=['origin_station_code', 'destination_station_code'])
    weekdayMat.FinalFactor.fillna(globalFactor, inplace=True)

    weekdayMat['MidweekDemand'] = weekdayMat.Demand * weekdayMat.FinalFactor

    return weekdayMat[['FromZone', 'ToZone', 'MidweekDemand']]

def applyGroupSplits(matrix, groupSplits):

    # Split out the matrix into where one end is a group and the other is travelcard; and where this is not the case
    groupTravelcard = matrix.loc[((matrix.FromZone=='XZA')&(matrix.ToZone.isin(groupSplits.OrigGroup.to_list()))) | ((matrix.ToZone=='XZA')&(matrix.FromZone.isin(groupSplits.OrigGroup.to_list())))].copy()
    matrixFil = matrix.loc[~((matrix.FromZone=='XZA')&(matrix.ToZone.isin(groupSplits.OrigGroup.to_list()))) & ~((matrix.ToZone=='XZA')&(matrix.FromZone.isin(groupSplits.OrigGroup.to_list())))].copy()

    print(matrix.MidweekDemand.sum())
    print(groupTravelcard.MidweekDemand.sum()+matrixFil.MidweekDemand.sum())

    # Process non-travelcard+group OD's
    groupSplits['Orig'] = np.where(groupSplits.OrigGroup=='', groupSplits.FromCRS, groupSplits.OrigGroup)
    groupSplits['Dest'] = np.where(groupSplits.DestGroup=='', groupSplits.ToCRS, groupSplits.DestGroup)

    splitMat = matrixFil.merge(groupSplits, how='outer', left_on=['FromZone', 'ToZone'], right_on=['Orig', 'Dest'])
    splitMat.Split.fillna(1, inplace=True)
    

    splitMat['Orig'] = np.where(splitMat.Orig.isna(), splitMat.FromZone, splitMat.Orig)
    splitMat['Dest'] = np.where(splitMat.Dest.isna(), splitMat.ToZone, splitMat.Dest)

    splitMat['Orig'] = np.where(splitMat.Orig == splitMat.OrigGroup, splitMat.FromCRS, splitMat.Orig)
    splitMat['Dest'] = np.where(splitMat.Dest == splitMat.DestGroup, splitMat.ToCRS, splitMat.Dest)
    
    splitMat['Demand'] = splitMat.MidweekDemand * splitMat.Split

    splitMat = splitMat[['Orig', 'Dest', 'Demand']]

    # process travelcard+group ODs
    groupTravelcard['JoinOrig'] = np.where(groupTravelcard.FromZone == 'XZA', 'XLD', groupTravelcard.FromZone)
    groupTravelcard['JoinDest'] = np.where(groupTravelcard.ToZone == 'XZA', 'XLD', groupTravelcard.ToZone)

    gtSplits = groupTravelcard.merge(groupSplits, how='left', left_on=['JoinOrig', 'JoinDest'], right_on=['Orig', 'Dest'])

    gtSplits.drop(['JoinOrig', 'JoinDest'], axis=1, inplace=True)

    gtSplits['Orig'] = np.where(gtSplits.FromZone == 'XZA', "XZA", gtSplits.FromCRS)
    gtSplits['Dest'] = np.where(gtSplits.ToZone == 'XZA', 'XZA', gtSplits.ToCRS)

    gtSplits['Demand'] = gtSplits.MidweekDemand * gtSplits.Split

    gtSplits = gtSplits[['Orig', 'Dest', 'Demand']]

    splitMat = pd.concat([splitMat, gtSplits], axis=0, ignore_index=True)
    
    print(matrix.MidweekDemand.sum())
    print(splitMat.Demand.sum())

    return splitMat

def applyTravelcardSplits(matrix, travelcardSplits):

    fromTC = matrix.loc[matrix.Orig == 'XZA'].copy()
    toTC = matrix.loc[matrix.Dest == 'XZA'].copy()
    noTC = matrix.loc[(matrix.Orig != 'XZA') & (matrix.Dest != 'XZA')].copy()

    test = fromTC.copy()
    fromTC.drop(['Orig'], axis=1, inplace=True)
    fromTC = fromTC.merge(travelcardSplits, left_on='Dest', right_on='Orig', how='left')
    fromTC.drop(['Orig'], axis=1, inplace=True)
    fromTC.rename({'Dest_x':'Dest', 'Dest_y':'Orig'}, axis=1, inplace=True)
    fromTC['Demand'] = fromTC.Demand*fromTC.Split
    fromTC = fromTC[['Orig', 'Dest', 'Demand']]

    toTC.drop(['Dest'], axis=1, inplace=True)
    toTC = toTC.merge(travelcardSplits, left_on='Orig', right_on='Orig', how='left')
    #toTC.rename({'D_x':'Dest', 'Dest_y':'Orig'}, axis=1, inplace=True)
    toTC['Demand'] = toTC.Demand*toTC.Split
    toTC = toTC[['Orig', 'Dest', 'Demand']]

    newMatrix = pd.concat([noTC, fromTC, toTC], ignore_index=True, axis=0)

    print(matrix.Demand.sum())
    print(newMatrix.Demand.sum())

    return newMatrix

def applyTimeProfiles(matrix, timeProfiles):

    timeProfiles = timeProfiles.groupby(['origin_station_code', 'destination_station_code'], as_index=False).sum()

    mergedMat = matrix.merge(timeProfiles, left_on=['Orig', 'Dest'], right_on=['origin_station_code', 'destination_station_code'], how='left')

    mergedMat.drop(['origin_station_code', 'destination_station_code'], axis=1, inplace=True)
    
    mergedMat['DailyTotal'] = 0

    # First, infill ODs with no time profile data with global splits

    totals = []
    for i in range(24):
        total = mergedMat[f'{i}'].sum()
        totals.append(total)
        mergedMat[f'{i}'].fillna(total, inplace=True)
        mergedMat['DailyTotal'] = mergedMat.DailyTotal + mergedMat[f'{i}']
        
    # Then find situations where time profiles are all 0 and infill these with globals also
    for i in range(24):
        mergedMat[f'{i}'] = np.where(mergedMat.DailyTotal == 0, totals[i], mergedMat[f'{i}'])

    for i in range(24):
        mergedMat['DailyTotal'] = np.where(mergedMat.DailyTotal == 0, sum(totals), mergedMat['DailyTotal'])

    # Then apply splits and create hourly matrices
    out_cols = ['Orig', 'Dest']
    mergedMat['Total'] =0
    mergedMat['DemandTotal'] =0
    for i in range(24):
        mergedMat[f'{i}'] = mergedMat[f'{i}']/mergedMat[f'DailyTotal']
        mergedMat[f'Matrix({i+1})'] = mergedMat.Demand * mergedMat[f'{i}']
        mergedMat['Total'] = mergedMat.Total + mergedMat[f'{i}']
        mergedMat['DemandTotal'] = mergedMat.DemandTotal + mergedMat[f'Matrix({i+1})']
        out_cols.append(f'Matrix({i+1})')

    
    mergedMat = mergedMat[out_cols].copy()

    return mergedMat



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


        weekdayMat = pd.read_csv(WeekdayMatrix, sep = '\t', skiprows= 12, low_memory = False, names=['FromZone', 'ToZone', 'Demand'], header=0, keep_default_na=False)
        weekdayMat = weekdayMat.loc[weekdayMat.FromZone != weekdayMat.ToZone].copy()
        weekdayMat.replace({'FromZone':dictCRS, 'ToZone':dictCRS}, inplace=True)

        midweekFactors = pd.read_csv(MidweekFactors, low_memory=False, keep_default_na=False)
        midweekFactors.replace({'origin_station_code':dictCRS, 'destination_station_code':dictCRS}, inplace=True)
        midweekMat = applyMidweekFactors(weekdayMat, midweekFactors, GlobalFactor)

        groupSplits = pd.read_csv(GroupedStationSplits, low_memory=False, keep_default_na=False)
        groupSplits.replace({'FromCRS':dictCRS, 'ToCRS':dictCRS}, inplace=True)
        groupSplits.Split = np.where(groupSplits.Split=="", 1, groupSplits.Split)
        groupSplits.Split = groupSplits.Split.astype(float)
        ungroupedMat = applyGroupSplits(midweekMat, groupSplits)

        travelcardSplits = pd.read_csv(TravelcardSplits, low_memory=False, keep_default_na=False)
        travelcardSplits.replace({'Orig':dictCRS, 'Dest':dictCRS}, inplace=True)
        splitMat = applyTravelcardSplits(ungroupedMat, travelcardSplits)

        dailyMatrix = splitMat.groupby(['Orig', 'Dest'], as_index=False).Demand.sum()

        logging.info(f'Daily demand total: {dailyMatrix.Demand.sum()}')
        print(dailyMatrix.Demand.sum())

        timeProfiles = pd.read_csv(TimeProfiles, low_memory=False, keep_default_na=False)
        timeProfiles.replace({'origin_station_code':dictCRS, 'destination_station_code':dictCRS}, inplace=True)

        hourlyMatrices = applyTimeProfiles(dailyMatrix, timeProfiles)
        total=0.0
        for i in range(24):
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
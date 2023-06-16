#%%

import glob
import math
import os
import pickle
import numpy as np
import pandas as pd
import win32com.client as com
from itertools import combinations
from matplotlib import pyplot as plt
from scipy.optimize import curve_fit
from scipy.stats import poisson

def poisson_func(x, lambda_):

    '''
    Defines a function from scipy.stats to be used in the scipy.optimimize curve_fit function.
        
        Parameters:
            x (float): k as per the scipy.stats.poisson documentation
            lamda_ (float): mu as per the scipy.stats.poisson documentation
    
        Returns:
            func: Cumulative Desity of the Poisson Distribution
    '''

    return poisson.cdf(x, lambda_)

def poissonFit(data):

    '''
    Calculates the mean assuming a poisson distribution of counts on a given day type at a given hour if the mean is unknown, providing the arithmetic mean otherwise.
        
        Parameters:
            data (pandas array-like): The data that is assume to be poisson distributed, with all unknown zalues > 0 AND < 10 replaced with 9
    
        Returns:
            mean: As calculated depending on whether unkown values were present or not
    '''
    
    #Check if 9 is in our data array. If not, we can skip the curve_fit process and just return the arithmetic mean, which will be both more accurate and significantly faster
    if 9 in data.values:
        
        #Define cumulative values, as we need to  fit to a CDF due to the missing values > 0 AND < 10
        cdf_array = data.value_counts().sort_index().cumsum()
        
        #This is the exception case when we have only numbers > 0 and < 10, There's not much we can do here but this seemed a sensible approach for consistency
        if cdf_array.size == 1:
            mean = curve_fit(poisson_func, [0, 9], [0, 1], bounds = (1, 9))[0][0]
        
        #Define our two arrays to apply in the curve_fit fuction
        else:
            counts = cdf_array.index
            cdf_vals = cdf_array.values/cdf_array.values[-1]

            #Try to apply the curve fit, with bounds to ensure we never get an impossible mean value in the event that Poisson dist is not a good fit to the observed data
            try:
                mean = curve_fit(poisson_func, counts, cdf_vals, bounds = (data.replace(9, 1).mean(), data.mean()))[0][0]

            #If a RuntimeError occurs, suggesting a response could not be found for a specific GroupBy array, then return None
            except RuntimeError:
                return None    
        
        #Return the final calculated value
        return mean
    
    #If there are no anonymized values in our array to begin with, we can simply return the arithmetic mean
    else:
        return data.mean()

def stationPlot(df, CRS, direction = None):

    '''
    Produces a plot of the process (N.B. Only for use in VS Code interactive mode).
        
        Parameters:
            df (pandas DataFrame): The DataFrame containing hourly matrix values to be plotted, with MultiIndex of the format (FROMZONE\CODE, TOZONE\CODE)
            CRS (string OR tuple of 2 strings): use a 3 letter CRS string to reference an individual station, or a tuple of from_CRS and to_CRS to plot a specific OD pair
            direction (None OR string): None if specific OD defined as CRS, Use 'from' OR 'to' flags otherwise (N.B. Not case sensitive)
    
        Returns:
            plt.bar: An hourly barplot of either from zone totals, to zone totals, or a specific OD
    '''
    
    #If direction == None, then we expect the tuple of from_CRS and to_CRS to be equivalent to the MultiIndex value we need, so we can simply use .loc
    if direction == None:
        toPlot = df.loc[CRS][range(24)]
    
    #Otherwise, we need to decide whether to sum on FROMZONE\CODE or TOZONE\CODE
    else:
        toPlot = df[df.index.get_level_values(f'{direction.upper()}ZONE\CODE') == CRS][range(24)].sum()
    
    #Plot the result
    plt.bar(toPlot.index,toPlot.values)

def processMND(path, connSharesDict):

    '''
    Processes the MND data if not already readable from the pickle in cached data.
        
        Parameters:
            path (os.path.dirname object OR string): The DataFrame containing hourly matrix values to be plotted, with MultiIndex of the format (FROMZONE\CODE, TOZONE\CODE)
            connSharesDict (dict): Defines connector shares to use for stations that are part of grouped ticketing areas. The user should ensure numbers add up to 1 for each group, and should ensure that XLD is included, and that XZA is NOT included (As XZA shares are assumed to be equal to XLD)
    
        Returns:
            finalMatrix: The demand matrix to be assigned imported into Visum
    '''
    
    #Define the pickle location to save to / read from
    myPickle = os.path.join(path, f'cached_data\\MND\\passenger_volume.p')
    
    #If the pickle exists, read from pickle
    if os.path.exists(myPickle):
        with open(myPickle, 'rb') as f:
            df_group, = pickle.load(f)
    
    #Otherwise, undertake the processing of the MND data
    else:
        
        #Get CSV files list from a folder & drop unused columns
        csv_files = glob.glob(os.path.join(path, 'input\\MND Hourly OD\\3_mid\\*.csv'))
        drop_cols = ['journey_time_50_pctl', 'journey_time_75_pctl', 'journey_time_95_pctl', 'journey_time_std', 'departure_wait_time_50_pctl',
                    'departure_wait_time_75_pctl', 'departure_wait_time_95_pctl', 'departure_wait_time_std', 'interchange_wait_time_50_pctl',
                    'interchange_wait_time_75_pctl', 'interchange_wait_time_95_pctl', 'interchange_wait_time_std',
                    'journey_time_mean', 'departure_wait_time_mean','interchange_wait_time_mean']

        #This creates a list of dataframes
        df_list = (pd.read_csv(file).drop(drop_cols, axis = 1) for file in csv_files)

        #Concatenate all DataFrames
        df = pd.concat(df_list, ignore_index = True)

        #Replace the anonymized data with the largest possible discrete value it could be for the CDF
        df.replace(4.42, 9, inplace = True)

        #Make int16 to queeze every second of speed we can out of embarrasing poissonFit
        df['passenger_volume'] = df['passenger_volume'].astype('int16')

        #Pivot to separate distinct hours into separate columns (as Poisson only valid assuming baseline rate of occurence is the same)
        #Then assume all entries completely missing from MND can be assumed to be a 0 count and fill in appropriately
        df_OD = df.pivot_table('passenger_volume', ['origin_station_code', 'destination_station_code', 'date'], 'hour').fillna(0).astype('int16').reset_index().drop('date', axis = 1)
        
        #Group by distinct OD pairs (again as Poisson only valid assuming baseline rate of occurence is the same), then apply the aggregation
        df_group = df_OD.groupby(['origin_station_code', 'destination_station_code']).agg(poissonFit)
        
        #Filter all rows and columns except those for which the aggregated GroupBy poissonFit approach failed
        df_validate = df_group[df_group.isna().any(axis = 1)].transpose()[df_group[df_group.isna().any(axis = 1)].transpose().isna().any(axis = 1)].isna()
        
        #Infill as many of the values that timed out as possible
        for i, row in df_validate.iterrows():
            for stationPair in df_validate.columns:
                if row[stationPair]:
                    array = df[(df['origin_station_code'] == stationPair[0]) & (df['destination_station_code'] == stationPair[1]) & (df['hour'] == i)]['passenger_volume']
                    df_group.at[stationPair, i] = poissonFit(array)

        #Write to pickle
        with open(myPickle, 'wb') as f:
            pickle.dump([df_group], f)

    #Discard internal trips from MND and any remaining OD pairs that still contain a nan value
    df_group_complete = df_group[(df_group.index.get_level_values(0) != df_group.index.get_level_values(1)) & (df_group.notna().all(axis = 1))]
    
    #Read the Daily matrix att file and rename the index of MND to ensure easy matching
    dailyMatrix = pd.read_csv(os.path.join(path, 'input\\Daily2022Matrix.att'), sep = '\t', skiprows= 12, low_memory = False).set_index(['FROMZONE\CODE', 'TOZONE\CODE'])
    df_group_complete.index.rename(dailyMatrix.index.names, inplace = True)

    #Use the connector shares to group into grouped ticketing areas also
    df_group_shared = df_group_complete.rename(index = connSharesDict)
    df_group_shared = df_group_shared.groupby(level = [0, 1]).sum()

    #Make a copy but this time apply it for the London TravelCard
    df_group_travelcard = df_group_shared.rename(index = {'XLD': 'XZA'})
    
    #Step 1: Apply the match to split Daily matrix by individual MND OD pair if possible 
    valsMNDinit = dailyMatrix[[]].join(df_group_complete, how = 'left')
    print(f'Note: {100*valsMNDinit.count().max()/len(valsMNDinit)}% of matrix cells split after consideration of unique MND OD pairs.')

    #Step 2: Otherwise, apply the match to split Daily matrix by individual MND OD if possible, also considering grouped Ticketing Areas
    valsMNDinit.update(df_group_shared, overwrite = False)
    print(f'Note: {100*valsMNDinit.count().max()/len(valsMNDinit)}% of matrix cells split after consideration of MND OD including grouped stations.')

    #Step 3: Otherwise, apply the match to split Daily matrix by individual MND OD if possible, also considering grouped London TravelCard
    valsMNDinit.update(df_group_travelcard, overwrite = False)
    print(f'Note: {100*valsMNDinit.count().max()/len(valsMNDinit)}% of matrix cells split after consideration of MND OD including London Travelcard.')

    #Step 4: Otherwise, apply the match to split Daily matrix by sum of matched MND row total and MND col total (Using just row or just col if only one end of the OD could be matched)
    valsMNDo = df_group_complete.groupby(level = 0).sum()
    valsMNDd = df_group_complete.groupby(level = 1).sum()
    valsMNDsum = valsMNDinit[valsMNDinit.isna().any(axis = 1)].add(valsMNDo, fill_value = 0).add(valsMNDd, fill_value = 0)
    valsMNDinter1 = valsMNDinit.add(valsMNDsum, fill_value = 0)
    print(f'Note: {100*valsMNDinter1.count().max()/len(valsMNDinter1)}% of matrix cells split after consideration of row / column sums of MND.')

    #Step 5: Otherwise, apply the match to split Daily matrix by sum of matched MND row total and MND col total, also considering grouped Ticketing Areas
    valsMNDoShared = df_group_shared.groupby(level = 0).sum()
    valsMNDdShared = df_group_shared.groupby(level = 1).sum()
    valsMNDsumShared = valsMNDinter1[valsMNDinter1.isna().any(axis = 1)].add(valsMNDoShared, fill_value = 0).add(valsMNDdShared, fill_value = 0)
    valsMNDinter2 = valsMNDinter1.add(valsMNDsumShared, fill_value = 0)
    print(f'Note: {100*valsMNDinter2.count().max()/len(valsMNDinter2)}% of matrix cells split after consideration of row / column sums of MND including grouped stations.')
    
    #Step 6: Otherwise, apply the match to split Daily matrix by sum of matched MND row total and MND col total, also considering Lonon TravelCard
    valsMNDoTravelcard = df_group_travelcard.groupby(level = 0).sum()
    valsMNDdTravelcard = df_group_travelcard.groupby(level = 1).sum()
    valsMNDsumTravelcard = valsMNDinter2[valsMNDinter2.isna().any(axis = 1)].add(valsMNDoTravelcard, fill_value = 0).add(valsMNDdTravelcard, fill_value = 0)
    valsMND = valsMNDinter2.add(valsMNDsumTravelcard, fill_value = 0)
    print(f'Note: {100*valsMND.count().max()/len(valsMND)}% of matrix cells split after consideration of row / column sums of MND including London Travelcard.')

    #Step 7: Otherwise, apply the match to split Daily matrix by considering the national aggregated sum of all MND data
    valsMNDnational = df_group_complete.sum()
    [valsMND[i].fillna(valsMNDnational[i], inplace = True) for i in range(24)]
    print(f'Note: {100*valsMND.count().max()/len(valsMND)}% of matrix cells split after consideration of MND national total.')

    #Calculate the hourly proportions for matched MND by dividing by the matched MND daily totals, then multiply the daily MOIRA matrix by these proportion factors 
    propsMND = valsMND.div(valsMND.sum(axis = 1), axis = 0)
    hourlyMatrix = propsMND.multiply(dailyMatrix[dailyMatrix.columns[-1]], axis = 0)

    #Check the totals of the daily & hourly MOIRA matrices, then join the results to the desired index and return the final result of the function
    print('Check Totals:')
    print(f'Daily Matrix Total: {dailyMatrix[dailyMatrix.columns[-1]].sum()}')
    print(f'Hourly Matrix Total: {hourlyMatrix.sum().sum()}')
    finalMatrix = dailyMatrix.join(hourlyMatrix, how = 'left')
    return finalMatrix

def main():

    path = os.path.dirname(__file__)
    connShares = pd.read_csv(os.path.join(path, 'input\\connector_shares.csv'), low_memory = False)
    connSharesDict = connShares.set_index('StationCRS')['CRS'].to_dict()
    myMatrix = processMND(path, connSharesDict).reset_index()
    fromZones = myMatrix[['$ODPAIR:FROMZONENO', 'FROMZONE\CODE']].rename(columns = {'$ODPAIR:FROMZONENO': 'ZONENO', 'FROMZONE\CODE': 'CODE'})
    toZones = myMatrix[['TOZONENO', 'TOZONE\CODE']].rename(columns = {'TOZONENO': 'ZONENO', 'TOZONE\CODE': 'CODE'})
    zones = pd.concat([fromZones.drop_duplicates(), toZones.drop_duplicates()]).drop_duplicates().set_index('ZONENO')

    Visum = com.Dispatch('Visum.Visum.230')
    Visum.IO.LoadVersion(os.path.join(path, 'output\\VISUM\\LOCs_and_PLTs_with_GTFS.ver'))
    
    allStopAreas = Visum.Net.StopAreas.FilteredBy(f'[NAME]="Platform Unknown"&[STOP\SUM:STOPAREAS\SUM:STOPPOINTS\COUNT:SERVINGVEHJOURNEYS]>0')
    atts = ['Code', 'NodeNo', 'XCoord', 'YCoord', 'Stop\\Name', 'Stop\\CRS']
    allStopAreasDF = pd.DataFrame(allStopAreas.GetMultipleAttributes(atts), columns = atts).set_index('Code')

    crsTIPLOCoverride = pd.read_csv(os.path.join(path, 'input\\CRS-TIPLOC_manual_override.csv'), low_memory = False)[['CRS', 'TIPLOC']].dropna().set_index('CRS')

    connSharesTravelcard = connShares[connShares['CRS'] == 'XLD'].copy()
    connSharesTravelcard.loc[connSharesTravelcard.index, 'CRS'] = 'XZA'
    connShares = pd.concat([connShares, connSharesTravelcard], ignore_index = True)

    Visum.Graphic.StopDrawing = True
    
    for i, row in zones.iterrows():
        aCRS = row['CODE']
        myConnShares = connShares[connShares['CRS'] == aCRS].merge(allStopAreasDF, 'left', left_on = 'StationCRS', right_on = 'Stop\\CRS')
        if len(myConnShares) > 0:
            for j, connRow in myConnShares.iterrows():
                if connRow['StationCRS'] in crsTIPLOCoverride.index:
                    myConnShares.loc[j, myConnShares.columns[-5:]]  = allStopAreasDF.loc[crsTIPLOCoverride.loc[connRow['StationCRS']]].iloc[0]
            weightedLoc = np.dot(myConnShares['ConnectorShare'], myConnShares[['XCoord', 'YCoord']])
            aZone = Visum.Net.AddZone(i, weightedLoc[0], weightedLoc[1])
            aZone.SetAttValue('Code', aCRS)
            aZone.SetAttValue('Name', myConnShares.loc[0, 'CRS_Name'])
            aZone.SetAttValue('SharePuT', True)
            for _, connRow in myConnShares.iterrows():
                try:
                    aConn = Visum.Net.AddConnector(aZone, connRow['NodeNo'])
                    aConn.SetAttValue('Weight(PuT)', 1000000*connRow['ConnectorShare'])
                    aConn.SetAttValue('ReverseConnector\\Weight(PuT)', 1000000*connRow['ConnectorShare'])
                    aConn.SetAttValue('T0_TSys(W)', 0)
                    aConn.SetAttValue('ReverseConnector\\T0_TSys(W)', 0)
                except:
                    print(f"Warning: No served node found for {connRow['StationCRS']}. Connector shares will be incorrect unless you define the desired CRS-TIPLOC match in the manual override csv.")
        else:
            if aCRS in crsTIPLOCoverride.index:
                myLoc = allStopAreasDF.loc[crsTIPLOCoverride.loc[aCRS, 'TIPLOC']]
            else:
                aStopArea = allStopAreasDF[allStopAreasDF['Stop\\CRS'] == aCRS]
                if aStopArea.shape[0] == 0:
                    print(f'Warning: No served node found for {aCRS}. Demand will be dropped unless you define the desired CRS-TIPLOC match in the manual override csv.')
                else:
                    if aStopArea.shape[0] > 1:
                        print(f'Warning: Multiple served nodes found for {aCRS}. The first match will be taken unless you define the desired CRS-TIPLOC match in the manual override csv.')
                    myLoc = aStopArea.iloc[0]
            aZone = Visum.Net.AddZone(i, myLoc['XCoord'], myLoc['YCoord'])
            aZone.SetAttValue('Code', aCRS)
            aZone.SetAttValue('Name', myLoc['Stop\\Name'])
            Visum.Net.AddConnector(aZone, myLoc['NodeNo'])

    LinkType = Visum.Net.AddLinkType(3)
    LinkType.SetAttValue('TSysSet', 'W')

    myCSV = pd.read_csv(os.path.join(path, 'input\\transfer_links.csv'), low_memory = False).set_index(['FromCRS', 'ToCRS'])
    
    for i, row in myCSV.iterrows():
        if i[0] < i[1]:
            if i[0] in crsTIPLOCoverride.index:
                myFromNode = allStopAreasDF.loc[crsTIPLOCoverride.loc[i[0]], 'NodeNo'][0]
                fromFlag = True
            else:
                try:
                    myFromNode = allStopAreasDF[allStopAreasDF['Stop\\CRS'] == i[0]]['NodeNo'][0]
                    fromFlag = True
                except:
                    print(f'Warning: No served node found for {i[0]}. No transfer link will be created unless you define the desired CRS-TIPLOC match in the manual override csv.')
                    fromFlag = False
            if i[1] in crsTIPLOCoverride.index:
                myToNode = allStopAreasDF.loc[crsTIPLOCoverride.loc[i[1]], 'NodeNo'][0]
                toFlag = True
            else:
                try:
                    myToNode = allStopAreasDF[allStopAreasDF['Stop\\CRS'] == i[1]]['NodeNo'][0]
                    toFlag = True
                except:
                    print(f'Warning: No served node found for {i[1]}. No transfer link will be created unless you define the desired CRS-TIPLOC match in the manual override csv.')
                    toFlag = False
            if fromFlag & toFlag:
                myLink = Visum.Net.AddLink(-1, myFromNode, myToNode, 3)
                myLink.SetAttValue('T_PUTSYS(W)', 60*row['TravelTime'])
                myLink.SetAttValue('REVERSELINK\\T_PUTSYS(W)', myCSV.loc[(i[1], i[0]), 'TravelTime'])

    LinkType = Visum.Net.AddLinkType(4)
    LinkType.SetAttValue('TSysSet', 'W')

    cc = list(combinations(allStopAreasDF[['NodeNo', 'XCoord', 'YCoord']].values, 2))

    for pair in cc:
        nodeFrom, xFrom, yFrom = pair[0]
        nodeTo, xTo, yTo = pair[1]
        distance = math.sqrt((xTo - xFrom)**2 + (yTo - yFrom)**2)
        if (distance < 250) & (nodeFrom < nodeTo):
            try:
                Visum.Net.AddLink(-1, nodeFrom, nodeTo, 4)
            except:
                print('Warning: This link has already been been manually defined. Therefore, no link is created.')
    
    Visum.Graphic.StopDrawing = False
    myIndex = pd.DataFrame(Visum.Net.ODPairs.GetMultipleAttributes(['FROMZONE\CODE', 'TOZONE\CODE']), columns = ['FROMZONE\CODE', 'TOZONE\CODE'])
    myExpandedMatrix = myIndex.merge(myMatrix, 'left', ['FROMZONE\CODE', 'TOZONE\CODE'])
    myTimeSeries = Visum.Net.AddTimeSeries(2, 1)
    for i in range(24):
        myVisumMatrix = Visum.Net.AddMatrix(100 + i, 2, 3)
        myHourlyMatrix = myExpandedMatrix[i].reset_index()
        myHourlyMatrix['index'] += 1
        Visum.Net.ODPairs.SetMultiAttValues(f'MatValue({str(100 + i)})', myHourlyMatrix.values)
        myTimeSeriesItem = myTimeSeries.AddTimeSeriesItem(3600*i, 3600*(i + 1))
        myTimeSeriesItem.SetAttValue('Matrix', f"Matrix([NO]={myVisumMatrix.AttValue('No')})")

    Visum.IO.SaveVersion(os.path.join(path, 'output\\VISUM\\LOCs_and_PLTs_with_GTFS_to_assign.ver'))

    print('Done')

# %%

if __name__ == "__main__":
    myMatrix = main()

# %%
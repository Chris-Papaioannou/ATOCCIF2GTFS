import glob
import pandas as pd
import os
import numpy as np
from itertools import product


path = os.path.dirname(__file__)

csv_files = glob.glob(os.path.join(path, 'input\\MND Hourly OD\\3_mid\\*.csv'))
drop_cols = ['journey_time_50_pctl', 'journey_time_75_pctl', 'journey_time_95_pctl', 'journey_time_std', 'departure_wait_time_50_pctl',
            'departure_wait_time_75_pctl', 'departure_wait_time_95_pctl', 'departure_wait_time_std', 'interchange_wait_time_50_pctl',
            'interchange_wait_time_75_pctl', 'interchange_wait_time_95_pctl', 'interchange_wait_time_std',
            'journey_time_mean', 'departure_wait_time_mean','interchange_wait_time_mean']

#This creates a list of dataframes
df_list = (pd.read_csv(file).drop(drop_cols, axis = 1) for file in csv_files)

#Concatenate all DataFrames
df = pd.concat(df_list, ignore_index = True)

df['Infilled'] = np.where(df.passenger_volume==4.42, True, False)

df_allDays = df.groupby(['origin_station_code', 'destination_station_code', 'hour', 'Infilled'], as_index=False).passenger_volume.count()
df_allDays.rename({'passenger_volume':'Datapoints'}, axis=1, inplace=True)

df_from = df_allDays.groupby(['origin_station_code', 'hour', 'Infilled'], as_index=False).Datapoints.sum()
#df_from = df_from.groupby(level=[0,1]).apply(lambda x: x / float(x.sum())).reset_index()
df_from.rename({'Datapoints':'FromTrips', 'origin_station_code':'Station'}, axis=1, inplace=True)

df_to = df_allDays.groupby(['destination_station_code', 'hour', 'Infilled'], as_index=False).Datapoints.sum()
#df_to = df_to.groupby(level=[0,1]).apply(lambda x: x / float(x.sum())).reset_index()
df_to.rename({'Datapoints':'ToTrips', 'destination_station_code':'Station'}, axis=1, inplace=True)

stations = df_from.Station.tolist()
stations.extend(df_to.Station.tolist())
all_stations = set(stations)

hours = list(range(24))
infills = [True, False]


df_all = pd.DataFrame(list(product(all_stations, hours, infills)), columns=['Station', 'hour', 'Infilled'])

df_from_to = df_all.merge(df_from.merge(df_to, how='outer'), how='outer').fillna(0).sort_values(['Station', 'hour', 'Infilled'])

df_from_to.to_csv("MND_Infills.csv", index=False)

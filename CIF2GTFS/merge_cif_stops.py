import os
import pandas as pd

path = os.path.dirname(__file__)

stops_att = pd.read_csv(os.path.join(path, 'input\\Stops.att'), sep = ';', header = 11, low_memory = False)

stops_att.rename(columns = {'$STOP:GTFS_STOP_ID': 'CRS'}, inplace = True)

cif_tiplocs = pd.read_csv(os.path.join(path, 'input\\cif_tiplocs.csv'), low_memory = False)

df = pd.merge(cif_tiplocs, stops_att, on = 'CRS')

df.to_csv(os.path.join(path, 'temp\\cif_tiplocs_loc.csv'), index = False)
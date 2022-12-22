import pandas as pd

stops_att = pd.read_csv('input\Stops.att', sep = ';', header = 11, low_memory = False)

stops_att['WKTLOCWGS84'] = stops_att['WKTLOCWGS84'].str.replace('POINT\(','', regex = True)
stops_att['WKTLOCWGS84'] = stops_att['WKTLOCWGS84'].str.replace('\)','', regex = True)

stops_att[['Lon', 'Lat']] = stops_att['WKTLOCWGS84'].str.split(' ', expand = True)
stops_att.drop(['WKTLOCWGS84'], axis = 1, inplace = True)

stops_att.rename(columns = {'$STOP:GTFS_STOP_ID': 'CRS'}, inplace = True)

cif_tiplocs = pd.read_csv('input\cif_tiplocs.csv', low_memory = False)

df = pd.merge(cif_tiplocs, stops_att, on = 'CRS')

df.to_csv('temp\cif_tiplocs_loc.csv', index = False)
import math
import os
import pandas as pd
from scipy.interpolate import interp1d
import win32com.client as com



path = os.path.dirname(__file__)

#TITLE BLOCK

df = pd.read_csv(os.path.join(path, 'temp\\stop_times_full.txt'), low_memory = False)

report = df['trip_id'].drop_duplicates(keep = False)

print(str(len(report)) + ' trips were dropped as they only had one stop. The following trip IDs were affected:')
print(report)

reduced_df = df[df.duplicated(subset = ['trip_id'], keep = False)]

reduced_df.to_csv(os.path.join(path, 'output\\stop_times.txt'), index = False)

#TITLE BLOCK

stop_ids = reduced_df['stop_id'].drop_duplicates().sort_values().values

attr_report = pd.read_csv(os.path.join(path, 'input\\NGD218 XYs Attribute 13 Report.csv'), low_memory = False)

attr_report_unique = attr_report[['ELR', 'Line', 'Structure Name', 'Station_Code']].drop_duplicates()
attr_report_unique.rename(columns = {'Station_Code': 'CRS'}, inplace = True)

cif_tiplocs = pd.read_csv(os.path.join(path, 'temp\\cif_tiplocs_loc.csv'), low_memory = False)
attr_report_unique = pd.merge(attr_report_unique, cif_tiplocs, on = 'CRS')

structures = [structure.split() for structure in attr_report_unique['Structure Name'].tolist()]

platforms = []
for structure in structures:
    try:
        platforms = platforms + [structure[structure.index('Platform') + 1]]
    except IndexError:
        platforms = platforms + ['NO WORD PROCEEDING PLATFORM']
    except ValueError:
        platforms = platforms + ['WORD PLATFORM IS NOT PRESENT']

attr_report_unique['Structure Name'] = attr_report_unique['Tiploc'] + '_' + platforms

Visum = com.Dispatch('Visum.Visum.220')
Visum.IO.LoadVersion(os.path.join(path, 'input\\DetailedNetwork.ver'))
Visum
station_prev = ''
station_no = 1
MyMapMatcher = Visum.Net.CreateMapMatcher()
for stop_id in stop_ids:
    id_split = stop_id.split('_')
    if id_split[0] != station_prev:
        station_prev = id_split[0]
        station_no += 1
        platform_no = 0
        s_No = 1000000 + 100 * station_no
        s_X = cif_tiplocs['XCOORD'][cif_tiplocs['Tiploc'].tolist().index(id_split[0])]
        s_Y = cif_tiplocs['YCOORD'][cif_tiplocs['Tiploc'].tolist().index(id_split[0])]
        Visum.Net.AddStop(s_No, s_X, s_Y)
        try:
            sa_Node = MyMapMatcher.GetNearestNode(s_X, s_Y, 250, False).Node
        except:
            sa_Link = MyMapMatcher.GetNearestLink(s_X, s_Y, 250, True)
            sa_Node = Visum.Net.AddNode(s_No, sa_Link.XPosOnLink, sa_Link.YPosOnLink)
            sa_Link.Link.SplitViaNode(sa_Node)
        Visum.Net.AddStopArea(s_No, s_No, sa_Node, s_X, s_Y)
    else:
        pass
    sp_No = s_No + platform_no
    unsatis = True
    alt = 0
    sp_Link = MyMapMatcher.GetNearestLink(s_X + dist * math.cos(ang * math.pi), s_Y + dist * math.sin(ang * math.pi), 250, True)
    sp = Visum.Net.AddStopPointOnLink(sp_No, s_No, sp_Link.Link.AttValue('FromNodeNo'), sp_Link.Link.AttValue('ToNodeNo'), True)
    while unsatis:
        try:
            RelPos = interp1d([0,1],[0 + alt, 1 - alt])
            sp.SetAttValue('RelPos', float(RelPos(sp_Link.RelPos)))
            unsatis = False
        except:
            alt += 0.001
    platform_no += 1
    print(sp_No)
    # if OSMmatch(id_split):
    #     initX = platform centroid
    #     initY = platform centroid
    # else:
    #     pass
    # if attr_match(id_split):
    #     apply TRIDlineFilt
    # elif attr_match_partial(id_split):
    #     apply TRIDlineFiltAlt
    # else:
    #     clear TRIDlineFilt
    # snap stop_id to network
    

print('Done')
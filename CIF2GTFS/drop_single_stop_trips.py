import os
import numpy as np
import pandas as pd
import re
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

attr_report_unique['Platform'] = platforms
attr_report_unique['AltPlatform'] = [re.sub('[^0-9]', '', platform) for platform in platforms]

def fix_dir_net(ver):
    Visum.IO.LoadVersion(ver)
    Links0 = Visum.Net.Links.GetFilteredSet('[TypeNo]=0')
    Links1 = Visum.Net.Links.GetFilteredSet('[TypeNo]=1')
    atts = ['OBJECTID', 'ASSETID', 'L_LINK_ID', 'L_SYSTEM', 'L_VAL', 'L_QUALITY', 'ELR', 'TRID',
            'TRCODE', 'L_M_FROM', 'L_M_TO', 'VERSION_NU', 'VERSION_DA', 'SOURCE', 'EDIT_STATU',
            'IDENTIFIED', 'TRACK_STAT', 'LAST_EDITE', 'LAST_EDI_1', 'CHECKED_BY', 'CHECKED_DA',
            'VALIDATED_', 'VALIDATED1', 'EDIT_NOTES', 'PROIRITY_A', 'SHAPE_LENG', 'TRID_CAT']
    Links0.SetMultipleAttributes(atts, Links1.GetMultipleAttributes(atts))
    Links0.GetFilteredSet('([TRCODE]>=10&[TRCODE]<=19)|([TRCODE]>=30&[TRCODE]<=39)|[TRCODE]>=50').SetAllAttValues('TypeNo', 1)
    Links1.GetFilteredSet('[TRCODE]>=10&[TRCODE]<=19').SetAllAttValues('TypeNo', 0)
    Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetAllAttValues('TSysSet', 'T')

def get_attr_report_fil(Tiploc_condit, df, platform):
    Visum.Net.Links.SetPassive()
    if np.any(Tiploc_condit & (df['Platform'] == platform)):
        df_fil = df[Tiploc_condit & (df['Platform'] == platform)]
    elif np.any(Tiploc_condit & (df['Platform'] == re.sub('[^0-9]', '', platform))):
        df_fil = df[Tiploc_condit & (df['Platform'] == re.sub('[^0-9]', '', platform))]
    elif np.any(Tiploc_condit & (df['AltPlatform'] == platform)):
        df_fil = df[Tiploc_condit & (df['AltPlatform'] == platform)]
    elif np.any(Tiploc_condit):
        df_fil = df[Tiploc_condit]
    else:
        df_fil = df[Tiploc_condit & (df['AltPlatform'] == re.sub('[^0-9]', '', platform))]
    if len(df_fil) > 0:
        for i in range(len(df_fil)):
            if i == 0:
                fil_string = f"([ELR]=\"{df_fil[f'ELR'].values[i]}\"&[TRID]=\"{df_fil['Line'].values[i]}\")"
            else:
                fil_string += f"|([ELR]=\"{df_fil[f'ELR'].values[i]}\"&[TRID]=\"{df_fil['Line'].values[i]}\")"
        fil_string = f'[TypeNo]=1&({fil_string})'
        Visum.Net.Links.GetFilteredSet(fil_string).SetActive()
    else:
        Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetActive()

def create_stop_point(X, Y, s_No, p_No):
    sp_No = s_No + p_No
    unsatis = True
    alt = [0, 0, 0, 0]
    sp_Link = MyMapMatcher.GetNearestLink(X, Y, 250, True, True)
    try:
        is_dir = sp_Link.Link.AttValue('ReverseLink\\TypeNo') == 0
    except:
        print(f'ERROR: NO FILTERED LINK WITHIN 250M FOR {stop_id} IN ATTRIBUTE REPORT. ATTRIBUTE REPORT IS PROBABLY INCORRECT...')
        Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetActive()
        is_dir = sp_Link.Link.AttValue('ReverseLink\\TypeNo') == 0
    try:
        sp = Visum.Net.AddStopPointOnLink(sp_No, s_No, sp_Link.Link.AttValue('FromNodeNo'), sp_Link.Link.AttValue('ToNodeNo'), is_dir)
    except:
        sp_Node = Visum.Net.AddNode(sp_No, sp_Link.Link.GetXCoordAtRelPos(0.5), sp_Link.Link.GetYCoordAtRelPos(0.5))
        sp_Link.Link.SplitViaNode(sp_Node)
        sp_Link = MyMapMatcher.GetNearestLink(X, Y, 250, True, True)
        sp = Visum.Net.AddStopPointOnLink(sp_No, s_No, sp_Link.Link.AttValue('FromNodeNo'), sp_Link.Link.AttValue('ToNodeNo'), is_dir)
    while unsatis:
        RelPos = interp1d([0, 0.5, 0.5, 1],[0 + alt[0], 0.5 - alt[1], 0.5 + alt[2], 1 - alt[3]])
        NewRelPos = float(RelPos(sp_Link.RelPos))
        shiftBool = [NewRelPos < 0.001, (NewRelPos > 0.499) & (NewRelPos <= 0.500), (NewRelPos >= 0.500) & (NewRelPos < 0.501), NewRelPos > 0.999]
        if np.any(shiftBool):
            alt = [altN + 0.001 if boolN else altN for altN, boolN in zip(alt, shiftBool)]
        else:
            try:
                sp.SetAttValue('RelPos', NewRelPos)
                unsatis = False
            except:
                alt = [altN + 0.001 for altN in alt]
    sp.SetAttValue('Code', stop_id)
    sp.SetAttValue('Name', f'Platform {id_split[1]}')

Visum = com.Dispatch('Visum.Visum.220')
fix_dir_net(os.path.join(path, 'input\\DetailedNetwork.ver'))
station_prev = ''
station_No = 1
MyMapMatcher = Visum.Net.CreateMapMatcher()
for stop_id in stop_ids[2224:]:
    id_split = stop_id.split('_')
    if id_split[0] != station_prev:
        station_prev = id_split[0]
        station_No += 1
        platform_No = 0
        stop_No = 1000000 + 100 * station_No
        s_X = cif_tiplocs['XCOORD'][cif_tiplocs['Tiploc'].tolist().index(id_split[0])]
        s_Y = cif_tiplocs['YCOORD'][cif_tiplocs['Tiploc'].tolist().index(id_split[0])]
        Visum.Net.AddStop(stop_No, s_X, s_Y)
        try:
            sa_Node = MyMapMatcher.GetNearestNode(s_X, s_Y, 250, False).Node
        except:
            sa_Link = MyMapMatcher.GetNearestLink(s_X, s_Y, 250, False)
            sa_Node = Visum.Net.AddNode(stop_No, sa_Link.XPosOnLink, sa_Link.YPosOnLink)
            sa_Link.Link.SplitViaNode(sa_Node)
        Visum.Net.AddStopArea(stop_No, stop_No, sa_Node, s_X, s_Y)
    else:
        pass
    get_attr_report_fil((attr_report_unique['Tiploc'] == id_split[0]), attr_report_unique, id_split[1])
    create_stop_point(s_X, s_Y, stop_No, platform_No)
    platform_No += 1
    # if OSMmatch(id_split):
    #     initX = platform centroid
    #     initY = platform centroid
    # else:
    #     pass

print('Done')
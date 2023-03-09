import os
os.environ['USE_PYGEOS'] = '0'
import overpy
import re
import time
import pickle
import wx
import numpy as np
import pandas as pd
import geopandas as gpd
import win32com.client as com
import matplotlib
matplotlib.use('Agg')
from matplotlib import pyplot as plt
from bng_latlon import OSGB36toWGS84, WGS84toOSGB36
from scipy.interpolate import interp1d
from shapely.geometry import Point, LineString, Polygon
from shapely.ops import nearest_points

def fix_dir_net(Visum, ver):

    #Load the version file
    Visum.IO.LoadVersion(ver)

    #Container object of reverse of links from directed shapefile
    Links0 = Visum.Net.Links.GetFilteredSet('[TypeNo]=0')

    #Container object of original links from directed shapefile
    Links1 = Visum.Net.Links.GetFilteredSet('[TypeNo]=1')

    #List of UDAs created upon import of directed shapefile
    atts = ['OBJECTID', 'ASSETID', 'L_LINK_ID', 'L_SYSTEM', 'L_VAL', 'L_QUALITY', 'ELR', 'TRID',
            'TRCODE', 'L_M_FROM', 'L_M_TO', 'VERSION_NU', 'VERSION_DA', 'SOURCE', 'EDIT_STATU',
            'IDENTIFIED', 'TRACK_STAT', 'LAST_EDITE', 'LAST_EDI_1', 'CHECKED_BY', 'CHECKED_DA',
            'VALIDATED_', 'VALIDATED1', 'EDIT_NOTES', 'PROIRITY_A', 'SHAPE_LENG', 'TRID_CAT']

    #Copy UDA values from original links to reverse links
    Links0.SetMultipleAttributes(atts, Links1.GetMultipleAttributes(atts))

    #Open reverse links if UP, BIDIRECT or TRCODE >= 50
    Links0.GetFilteredSet('([TRCODE]>=10&[TRCODE]<=19)|([TRCODE]>=30&[TRCODE]<=39)|[TRCODE]>=50').SetAllAttValues('TypeNo', 1)

    #Close original links if UP
    Links1.GetFilteredSet('[TRCODE]>=10&[TRCODE]<=19').SetAllAttValues('TypeNo', 0)

    #Set TSys for open links
    Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetAllAttValues('TSysSet', '2')

def overpass_query(overpassQLstring):
    
    #Create API object and boolean switch
    apiPy = overpy.Overpass()
    unsatis = True
    
    #Keep trying to access the API until successful
    while unsatis:
        try:
            result = apiPy.query(overpassQLstring)
            unsatis = False
        
        #N.B. This except is generic, so will go into an infinite loop if internet connection is down or if overpassQLstring is invalid format
        except:
            time.sleep(2)
            pass
    
    #Return the API query result
    return result

def get_OSM_node(crs, desc):
    
    #This function will only be run if a pickle file is not already present, so warning message are also saved to pickle for future reference
    OSMnodeWarning = ''
    
    #Look for nodes in OSM API with the correct 3 Letter string for tag ref:crs
    myOSMnode = overpass_query(f'node["ref:crs"~"^{crs}$",i];out;')
    
    #Check if API result is blank and rerun again with less stringent controls if necessary
    if len(myOSMnode.nodes) == 0:
        OSMnodeWarning = f'WARNING: {crs} - {desc}, No node with ref:crs ~ {crs} found in OSM. Searching for (railway = station) & (name ~ {desc}) instead.'
        myOSMnode = overpass_query(f'node["railway"="station"]["name"~"^{desc}$",i];out;')
        if len(myOSMnode.nodes) == 0:
            myOSMnode = overpass_query(f'node["railway"="station"]["alt_name"~"^{desc}$",i];out;')
            if len(myOSMnode.nodes) == 0:
                OSMnodeWarning = f'ERROR: No node with (ref:crs ~ {crs}) OR ((railway = station) & (name / alt_name ~ {desc})) found in OSM.'
        else:
            if len(myOSMnode.nodes) > 1:
                OSMnodeWarning = f'WARNING: {crs} - {desc}, There is > 1 node with (railway = station) & (name / alt_name ~ {desc}) in OSM. The first instance is taken.'
    
    #Otherwise check if there was more than one OSM node returned in the first instance
    else:
        if len(myOSMnode.nodes) > 1:
            OSMnodeWarning = f'WARNING: {crs} - {desc}, There is > 1 node with ref:crs ~ {crs} in OSM. The first instance is taken.'
    
    #Define coordinate of OSM node as a shapely Pont in BNG format and return alongside any warning
    EastNorth = Point(WGS84toOSGB36(float(myOSMnode.nodes[0].lat), float(myOSMnode.nodes[0].lon)))
    return EastNorth, OSMnodeWarning

def str_clean(myStr, desc):
    
    #Make string upper case, and if the platform name simply contains the station name, treat the same as blank or missing OSM tags
    myStr = myStr.upper()
    if myStr == desc:
        return 'REF ERROR'
    
    #Replace & and : with ; (the most common char used for df line duplication) and tidy string
    else:
        replace = ['&', ':']
        for rep in replace:
            myStr = myStr.replace(rep, ';')
        remove = ['PLATFORMS', 'PLATFORM', ' ']
        for rem in remove:
            myStr = myStr.replace(rem, '')
        return myStr

def process_platformWays(myPlatformWays, crs, desc, EastNorth):
    
    #Close previous plot(s), then pre-define new plot with equal axis scales and set title
    plt.close('all')
    fig, ax2 = plt.subplots(figsize = (20, 16))
    ax2.set_aspect('equal', adjustable = 'box')
    ax2.set_title(f'{crs}: {desc}')
    
    #Iterate through ways in OSM query result
    for way in myPlatformWays:
        
        #Project nodes in way to BNG
        try:
            way.bngs = np.array([WGS84toOSGB36(float(node.lat), float(node.lon)) for node in way.nodes])
        
        #If there is an error, this is likely because the way is actually a multipolygon relation with holes that we appended earlier
        except:
            
            #We are not interested in the holes (i.e. inner faces), so we take the nodes from the first outer face in the object 
            way.nodes = way._result._ways[way.members[[member.role == 'outer' for member in way.members].index(True)].ref].nodes
            way.bngs = np.array([WGS84toOSGB36(float(node.lat), float(node.lon)) for node in way.nodes])
        
        #Create a boolean to check if first and last node do not close the shape and then create the corresponding Shapely object accordingly
        isLine = any(way.bngs[0] != way.bngs[-1])
        if isLine:
            shape = LineString(way.bngs)
        else:
            shape = Polygon(way.bngs)
        
        #Cast as a geopandas object and add to our pre-defined plot and get the minimum rotated bounding rectangle
        gds = gpd.GeoSeries(shape)
        gds.plot(edgecolor = 'black', color = 'lightcoral', ax = ax2)
        mrr = shape.minimum_rotated_rectangle

        #If area of minimum bounding rectangle is 0, this means object is straight line, so return shapely centroid in the normal fashion
        if mrr.area == 0:
            way.bng = mrr.centroid
        
        #Otherwise, split the bounding box into 4 edges, creating a line linking the midpoints of the 1st and 2nd edges sorted by length
        else:
            coords = [c for c in mrr.boundary.coords]
            mrr_long = sorted([LineString([c1, c2]) for c1, c2 in zip(coords, coords[1:])], key = lambda x: x.length, reverse = True)[:2]
            mrr_bisect = LineString([mrr_long[0].centroid, mrr_long[1].centroid])
            
            #Use this bisection line to calculate the centroid of the intersection (nearest points used due to rounding error issues)
            if isLine:
                way.bng = nearest_points(shape, mrr_bisect.centroid)[0]
            else:
                way.bng = gpd.clip(gpd.GeoSeries(mrr_bisect), gds).centroid.values[0]
        
        #Calculate the distance from the OSM node and plot the platform centroid locations
        way.dist = way.bng.distance(EastNorth)
        platform_centroid = gpd.GeoSeries(way.bng)
        platform_centroid.plot(color = 'blue', ax = ax2)
        
        #Clean the OSM tag strings (using local_ref followed by name if default ref not found) and annotate the platform centroids
        try:
            way.tags['ref'] = str_clean(way.tags['ref'], desc)
        except:
            try:
                way.tags['ref'] = str_clean(way.tags['local_ref'], desc)
            except:
                try:
                    way.tags['ref'] = str_clean(way.tags['name'], desc)
                except:
                    way.tags['ref'] = 'REF ERROR'
        plt.annotate(' ' + way.tags['ref'], (way.bng.x, way.bng.y))
    
    #Return the now processed platform information and the figure
    return myPlatformWays, fig

def get_OSM_platform_data(path, crs, desc, bound):
    
    #Check if OSM data has been saved as pickle already for this station, and if so, read from pickle file, otherwise query them from OSM
    myPickle = os.path.join(path, f'OSM_pickles\\{crs}_platforms.p')
    if os.path.exists(myPickle):
        with open(myPickle, 'rb') as f:
            EastNorth, OSMnodeWarning, platformWays = pickle.load(f)
    else:
        EastNorth, OSMnodeWarning = get_OSM_node(crs, desc)
        min = OSGB36toWGS84(EastNorth.x - bound, EastNorth.y - bound)
        max = OSGB36toWGS84(EastNorth.x + bound, EastNorth.y + bound)
        platformWays = overpass_query(f'way({min[0]},{min[1]},{max[0]},{max[1]})["railway"~"platform"];(._;>;);out body;').ways
        platformRelations = overpass_query(f'relation({min[0]},{min[1]},{max[0]},{max[1]})["railway"~"platform"];(._;>;);out body;').relations
        platformWays = platformWays + platformRelations
        with open(myPickle, 'wb') as f:
            pickle.dump([EastNorth, OSMnodeWarning, platformWays], f)
    
    #If present, print the OSM node warning preserved in the pickle file from the originally run API request
    if OSMnodeWarning != '':
        print(OSMnodeWarning)
    
    #If no platform ways are present in the OSM data, return just the OSM node location and an empty list for platforms
    if len(platformWays) == 0:
        OSMstation = {'Location': EastNorth, 'Platforms': []}
        return OSMstation
    
    #Process the OSM platform way data and save the resultant figure as a png file
    platformWays, myFig = process_platformWays(platformWays, crs, desc, EastNorth)
    myFig.savefig(os.path.join(path, f'OSM_images\\{crs}_platforms.png'))
    
    #Convert the processed OSM platform way data into a pandas DataFrame and return alongside OSM station node location
    c1 = [way.tags['ref'] for way in platformWays]
    c2 = [way.bng for way in platformWays]
    c3 = [way.dist for way in platformWays]
    myCols = {'Platform': c1, 'Location': c2, 'Dist': c3}
    dfPlatforms = pd.DataFrame.from_dict(myCols)
    dfPlatforms = dfPlatforms.explode('Platform').reset_index(drop = True)
    dfPlatforms = dfPlatforms.join(dfPlatforms.pop('Platform').str.split(';', expand = True))
    dfPlatforms = dfPlatforms.melt(dfPlatforms.columns[:len(myCols)-1], dfPlatforms.columns[len(myCols)-1:])
    dfPlatforms = dfPlatforms.rename(columns = {'value': 'Platform'}).drop('variable', 1).sort_values('Dist').reset_index(drop = True)
    OSMstation = {'Location': EastNorth, 'Platforms': dfPlatforms}
    return OSMstation

def get_attr_report_fil(Tiploc_condit, df, platform, pNumer):
    if np.any(Tiploc_condit & (df['Platform'] == platform)):
        df_fil = df[Tiploc_condit & (df['Platform'] == platform)]
    elif np.any(Tiploc_condit & (df['Platform'] == pNumer)):
        df_fil = df[Tiploc_condit & (df['Platform'] == pNumer)]
    elif np.any(Tiploc_condit & (df['AltPlatform'] == platform)):
        df_fil = df[Tiploc_condit & (df['AltPlatform'] == platform)]
    elif np.any(Tiploc_condit & (df['AltPlatform'] == pNumer)):
        df_fil = df[Tiploc_condit & (df['AltPlatform'] == pNumer)]
    else:
        df_fil = df[Tiploc_condit]
    if len(df_fil) > 0:
        for i in range(len(df_fil)):
            if i == 0:
                fil_string = f"([ELR]=\"{df_fil[f'ELR'].values[i]}\"&[TRID]=\"{df_fil['Line'].values[i]}\")"
            else:
                fil_string += f"|([ELR]=\"{df_fil[f'ELR'].values[i]}\"&[TRID]=\"{df_fil['Line'].values[i]}\")"
        fil_string = f'[TypeNo]=1&({fil_string})'
        return fil_string
    else:
        fil_string = '[TypeNo]=1'
        return fil_string

def get_sp_No(s_No, pNumer, pAlpha):
    if pNumer == '':
        numerNo = 280
    elif int(pNumer) < 28:
        numerNo = 10*int(pNumer)
    else:
        print(f'ERROR: Unexpected platform number {pNumer}, platform numbers 0 to 27 supported.')
    alphaIndex1 = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'N', 'W', 'S']
    alphaIndex2 = ['L', 'M', 'R', 'X', 'DF', 'DM', 'DW', 'UB', 'UF', 'UM']
    alphaIndex3 = ['FL', 'TL',
                    'PLACEHOLDER (2 OR 3 CHAR)', 'PLACEHOLDER (2 OR 3 CHAR)', 'PLACEHOLDER (2 OR 3 CHAR)', 'PLACEHOLDER (2 OR 3 CHAR)',
                    'PLACEHOLDER (2 OR 3 CHAR)', 'PLACEHOLDER (2 OR 3 CHAR)', 'PLACEHOLDER (2 OR 3 CHAR)', 'PLACEHOLDER (2 OR 3 CHAR)',
                    'BAY', 'DFL', 'DML', 'DPL', 'SGL', 'UFL', 'UML', 'UPL', 'URL', 'PLACEHOLDER (3 CHAR ONLY)']
    if max(len(alphaIndex1), len(alphaIndex2)) > 10:
        print('ERROR: AlphaIndex1 & AlphaIndex2 should represent a number from 0 to 9.')
    if len(alphaIndex3) > 20:
        print('ERROR: AlphaIndex3 should represent a number from 0 to 19.')
    if pAlpha in alphaIndex1:
        alphaNo = alphaIndex1.index(pAlpha)
    elif pAlpha in alphaIndex2:
        alphaNo = 300 + alphaIndex2.index(pAlpha)
    elif pAlpha in alphaIndex3:
        alphaNo = 600 + alphaIndex3.index(pAlpha)
    else:
        print(f'ERROR: Unexpected pAlpha format. If needed, replace one of the placeholder locations in alphaIndex.')
    return 1000*s_No + numerNo + alphaNo

def create_stop_point(Visum, X, Y, s_No, bound, crs, desc, platform, pNumer, pAlpha):
    sp_No = get_sp_No(s_No, pNumer, pAlpha)
    if platform == '':
        sa = Visum.Net.AddStopArea(sp_No, s_No, s_No, X, Y)
        sa.SetAttValue('Code', f'{crs}_')
        sa.SetAttValue('Name', 'Platform Unknown')
        sp = Visum.Net.AddStopPointOnNode(sp_No, sa, s_No)
        sp.SetAttValue('Code', f'{crs}_')
        sp.SetAttValue('Name', 'Platform Unknown')
    else:
        unsatis = True
        alt = [0, 0, 0, 0]
        MyMapMatcher = Visum.Net.CreateMapMatcher()
        sp_Link = MyMapMatcher.GetNearestLink(X, Y, bound, True, True)
        try:
            is_dir = sp_Link.Link.AttValue('ReverseLink\\TypeNo') == 0
        except:
            print(f'WARNING: No filtered link within {bound}m for {crs} - {desc}, Platform {platform}.')
            print('         Attribute Report Filter is probably incorrect. Trying again without filter active.')
            Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetActive()
            sp_Link = MyMapMatcher.GetNearestLink(X, Y, bound, True, True)
            is_dir = sp_Link.Link.AttValue('ReverseLink\\TypeNo') == 0
        if sp_Link.RelPos < 0.5:
            sa_Node = sp_Link.Link.AttValue('FromNodeNo')
        else:
            sa_Node = sp_Link.Link.AttValue('ToNodeNo')
        sa = Visum.Net.AddStopArea(sp_No, s_No, sa_Node, sp_Link.XPosOnLink, sp_Link.YPosOnLink)
        sa.SetAttValue('Code', f'{crs}_{platform}')
        sa.SetAttValue('Name', f'Platform {platform}')
        try:
            sp = Visum.Net.AddStopPointOnLink(sp_No, sa, sp_Link.Link.AttValue('FromNodeNo'), sp_Link.Link.AttValue('ToNodeNo'), is_dir)
        except:
            sp_Node = Visum.Net.AddNode(sp_No, sp_Link.Link.GetXCoordAtRelPos(0.5), sp_Link.Link.GetYCoordAtRelPos(0.5))
            Visum.Net.StopAreas.ItemByKey(sp_No).SetAttValue('NodeNo', sp_No)
            sp_Link.Link.SplitViaNode(sp_Node)
            sp_Link = MyMapMatcher.GetNearestLink(X, Y, bound, True, True)
            sp = Visum.Net.AddStopPointOnLink(sp_No, sa, sp_Link.Link.AttValue('FromNodeNo'), sp_Link.Link.AttValue('ToNodeNo'), is_dir)
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
        sp.SetAttValue('Code', f'{crs}_{platform}')
        sp.SetAttValue('Name', f'Platform {platform}')

def get_platform_loc(platform, alt_platform, s_loc, CRSplatforms, crs, desc):
    if platform == '':
        platformLocation = s_loc
    elif len(CRSplatforms) == 0:
        platformLocation = s_loc
        print(f'WARNING: {crs} - {desc} has no OSM platforms within bound. Used OSM station node location for Platform {platform} instead.')
    else:
        try:
            try:
                my_index = CRSplatforms['Platform'].tolist().index(platform)
            except:
                alt_platform
                try:
                    my_index = CRSplatforms['Platform'].tolist().index(alt_platform)
                    print(f'WARNING: {crs} - {desc}, Platform {platform} is not in OSM. Matched to {alt_platform} instead.')
                except:
                    try:
                        my_index = [re.sub('[^0-9]', '', ref) for ref in CRSplatforms['Platform'].tolist()].index(alt_platform)
                        alt_platform_2 = CRSplatforms['Platform'][my_index]
                        print(f'WARNING: {crs} - {desc}, Platform {platform} is not in OSM. Matched to {alt_platform_2} instead.')
                    except:
                        my_index = CRSplatforms['Platform'].tolist().index('REF ERROR')
                        print(f'WARNING: {crs} - {desc}, Platform {platform} is not in OSM. Matched to nearest unknown platform.')
            platformLocation = CRSplatforms['Location'][my_index]
        except:
            platformLocation = s_loc
            if platform not in ['DM', 'DPL', 'UM', 'UPL']:
                print(f'WARNING: {crs} - {desc}, Platform {platform} is not in OSM & no unknown platform available. Used OSM station node location instead.')
            else:
                pass
                #N.B. DO something smart with the line filters for Up / Down platform? Maybe after speaking with NR?
    return platformLocation

def main(skipped_rows = 0):

    #Define path and read oringinal stop times gtfs output from C# CIF to GTFS process
    path = os.path.dirname(__file__)
    df = pd.read_csv(os.path.join(path, 'temp\\stop_times_full.txt'), low_memory = False)

    #Get pandas DataFrame of unique trip IDs (i.e. single stop trips) only for reporting
    report = df['trip_id'].drop_duplicates(keep = False)
    print(f'WARNING: {len(report)} trips were dropped as they only had one stop. The following trip IDs were affected:')
    print(report)

    #Drop unique trip IDs (i.e. single stop trips) and output to the final location, and get a unique list of stop IDs
    reduced_df = df[df.duplicated(subset = ['trip_id'], keep = False)].reset_index(drop = True)
    reduced_df_IDfix = reduced_df.copy()
    for i, stop_id in enumerate(reduced_df_IDfix['stop_id']):
        crs, platform = stop_id.split('_')
        pNumer = re.sub('[^0-9]', '', platform)
        pAlpha = re.sub('[^A-Z]', '', platform)
        s_No = 10000*(ord(crs[0]) - 55) + 100*(ord(crs[1]) - 55) + (ord(crs[2]) - 55)
        reduced_df_IDfix.loc[i, 'stop_id'] = get_sp_No(s_No, pNumer, pAlpha)
    reduced_df_IDfix.to_csv(os.path.join(path, 'output_GTFS\\stop_times.txt'), index = False)
    stop_ids = reduced_df['stop_id'].drop_duplicates().sort_values().values

    #Read attribute report file and get a unique list of ELR, Line, Structure Name, and Station Code combinations
    attr_report = pd.read_csv(os.path.join(path, 'input\\NGD218 XYs Attribute 13 Report.csv'), low_memory = False)
    attr_report_unique = attr_report[['ELR', 'Line', 'Structure Name', 'Station_Code']].drop_duplicates()
    attr_report_unique.rename(columns = {'Station_Code': 'CRS'}, inplace = True)

    cif_tiplocs = pd.read_csv(os.path.join(path, 'input\\cif_tiplocs.csv'), low_memory = False)

    #N.B. This is a bit of a mess but wouldn't bother changing things until a proper conversation about attribute report file with NR.
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

    #The actual main process
    Visum = com.Dispatch('Visum.Visum.230')
    fix_dir_net(Visum, os.path.join(path, 'input\\DetailedNetwork.ver'))
    MyMapMatcher = Visum.Net.CreateMapMatcher()
    station_prev = ''

    #Define Progress Bar
    class ProgWin(wx.Frame): 

        def __init__(self, parent, title): 
            super(ProgWin, self).__init__(parent, title = title,size = (300, 200))  
            self.InitUI() 
                
        def InitUI(self):    
            self.count = 0 
            pnl = wx.Panel(self)
                
            self.gauge = wx.Gauge(pnl, range = len(stop_ids[skipped_rows:]), size = (300, 25), style =  wx.GA_HORIZONTAL) 
                
            self.SetSize((300, 100)) 
            self.Centre() 
            self.Show(True)
                        
    ex = wx.App() 
    prog = ProgWin(None, 'wx.Gauge')

    for progNo, stop_id in enumerate(stop_ids[skipped_rows:]):
        crs, platform = stop_id.split('_')
        pNumer = re.sub('[^0-9]', '', platform)
        pAlpha = re.sub('[^A-Z]', '', platform)
        if crs != station_prev:
            station_prev = crs
            myDesc = cif_tiplocs['Description'][cif_tiplocs['CRS'].tolist().index(crs)]
            myCRSno = 10000*(ord(crs[0]) - 55) + 100*(ord(crs[1]) - 55) + (ord(crs[2]) - 55)
            myCRSdata = get_OSM_platform_data(path, crs, myDesc, 250)
            sLoc = myCRSdata['Location']
            s = Visum.Net.AddStop(myCRSno, sLoc.x, sLoc.y)
            s.SetAttValue('Code', crs)
            s.SetAttValue('Name', myDesc)
            Visum.Net.AddNode(myCRSno, sLoc.x, sLoc.y)
            unsatis = True
            fil_string = '[TYPENO]=1'
            nTRID = 1
            while unsatis & (nTRID <= 33):
                Visum.Net.Links.SetPassive()
                Visum.Net.Links.GetFilteredSet(fil_string).SetActive()
                split_Link = MyMapMatcher.GetNearestLink(sLoc.x, sLoc.y, 250, True, True)
                unsatis = split_Link.Success
                if unsatis:
                    split_TRID = split_Link.Link.AttValue('TRID')
                    split_no = 100*myCRSno + nTRID
                    split_Node = Visum.Net.AddNode(split_no, split_Link.XPosOnLink, split_Link.YPosOnLink)
                    split_Link.Link.SplitViaNode(split_Node)
                    Visum.Net.Links.ItemByKey(split_Link.Link.AttValue('FROMNODENO'), split_no).SetNo(split_no + 33)
                    Visum.Net.Links.ItemByKey(split_no, split_Link.Link.AttValue('TONODENO')).SetNo(split_no + 66)
                    try:
                        Visum.Net.AddLink(split_no, split_no, myCRSno, 2)
                    except:
                        print('break me')
                    fil_string += f"&[TRID]!=\"{split_TRID}\""
                    nTRID += 1

            turn_fil_string = f'([VIANODENO]={str(myCRSno)}&[TYPENO]!=4)|([FROMNODENO]={str(myCRSno)}&[TYPENO]=4)|([TONODENO]={str(myCRSno)}&[TYPENO]=4)'
            Visum.Net.Turns.FilteredBy(turn_fil_string).SetAllAttValues('TSysSet', '')
        Visum.Net.Links.SetPassive()
        my_fil_string = get_attr_report_fil((attr_report_unique['CRS'] == crs), attr_report_unique, platform, pNumer)
        Visum.Net.Links.GetFilteredSet(my_fil_string).SetActive()
        platformLocation = get_platform_loc(platform, pNumer, sLoc, myCRSdata['Platforms'], crs, myDesc)
        create_stop_point(Visum, platformLocation.x, platformLocation.y, myCRSno, 250, crs, myDesc, platform, pNumer, pAlpha)
        prog.gauge.SetValue(progNo)
    
    DFcols_Visum = ['No', 'Code', 'Name', 'YCoord', 'XCoord']
    DFcols_GTFS = ['stop_id', 'stop_code', 'stop_name', 'stop_lat', 'stop_lon']
    DF_s = pd.DataFrame(Visum.Net.Stops.GetMultipleAttributes(DFcols_Visum), columns = DFcols_GTFS)
    DF_s['location_type'] = 1
    DF_sp = pd.DataFrame(Visum.Net.StopPoints.GetMultipleAttributes(DFcols_Visum + ['StopArea\\StopNo']), columns = DFcols_GTFS + ['parent_station'])
    DF_sp['location_type'] = 0
    DF = pd.concat([DF_s, DF_sp], axis = 0)
    DF['stop_id'] = pd.to_numeric(DF['stop_id'],errors='coerce').astype('Int64')
    DF['parent_station'] = pd.to_numeric(DF['parent_station'],errors='coerce').astype('Int64')
    DF.to_csv(os.path.join(path, 'output_GTFS\\stops.txt'), index = False)
    Visum.IO.SaveVersion(os.path.join(path, 'output_Visum\\DetailedNetwork_Processed.ver'))
    print('done')

if __name__ == "__main__":
    main()
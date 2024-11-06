import sys
import pathlib

cur_dir = pathlib.Path(__file__).parent.resolve()
sys.path.append(f"{cur_dir}\\src")

import json
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
from itertools import combinations
import math
import datetime
import traceback

sys.path.append(os.path.dirname(__file__))

import get_inputs as gi

def fixDirectedNet(Visum, reversedELRs, TSysDefs, railBased, PTpermitted):

    '''
    PLACEHOLDER.
        
        Parameters:
            Visum (os object): PLACEHOLDER
            reversedELRs (): PLACEHOLDER
            TSysDefs (): PLACEHOLDER
            railBased (): PLACEHOLDER
            PTpermitted (): PLACEHOLDER
    
        Returns:
            None
    '''

    #Container object of reverse of links from directed shapefile
    Links0 = Visum.Net.Links.GetFilteredSet('[TypeNo]=0')

    #Container object of original links from directed shapefile
    Links1 = Visum.Net.Links.GetFilteredSet('[TypeNo]=1')

    #Create Boolean UDAs for coorecting the directionality of the shapefile when down is not the open direction as expected
    UDAs = ['is_CLOSED', 'is_DOWN', 'is_UP', 'is_BIDIRECT', 'is_REVERSE']
    for uda in UDAs:
        Visum.Net.Links.AddUserDefinedAttribute(uda, uda, uda, 9)
    
    #Set boolean values for DOWN, UP, and BIDIRECT respectively (for now assume LOOP can be treated the same as DOWN)
    Links1.GetFilteredSet('WORDN([TRACK_STAT],"DISCONNECTED",1)!=[TRACK_STAT]|WORDN([TRACK_STAT],"UNUSED",1)!=[TRACK_STAT]').SetAllAttValues('is_CLOSED', True)
    Links1.GetFilteredSet('([TRCODE]>=20&[TRCODE]<=29)|([TRCODE]>=40&[TRCODE]<=49)').SetAllAttValues('is_DOWN', True)
    Links1.GetFilteredSet('[TRCODE]>=10&[TRCODE]<=19').SetAllAttValues('is_UP', True)
    Links1.GetFilteredSet('([TRCODE]>=30&[TRCODE]<=39)|[TRCODE]>=50').SetAllAttValues('is_BIDIRECT', True)
    
    #Iterate through Shapefile directed links and determine whether to reverse them or not
    iterateELRs = Links1.GetMultipleAttributes(['ELR', 'is_DOWN', 'is_UP'])
    isReversed = [[link[1]] if link[0] in reversedELRs else [link[2]] for link in iterateELRs]
    Links1.SetMultipleAttributes(['is_REVERSE'], isReversed)

    #List of UDAs created upon import of directed shapefile
    atts = ['OBJECTID', 'ASSETID', 'L_LINK_ID', 'L_SYSTEM', 'L_VAL', 'L_QUALITY', 'ELR', 'TRID',
            'TRCODE', 'L_M_FROM', 'L_M_TO', 'VERSION_NU', 'VERSION_DA', 'SOURCE', 'EDIT_STATU',
            'IDENTIFIED', 'TRACK_STAT', 'LAST_EDITE', 'LAST_EDI_1', 'CHECKED_BY', 'CHECKED_DA',
            'VALIDATED_', 'VALIDATED1', 'EDIT_NOTES', 'PROIRITY_A', 'SHAPE_LENG', 'TRID_CAT',
            'is_DOWN', 'is_UP', 'is_BIDIRECT', 'is_REVERSE']

    #Copy UDA values from original links to reverse links
    Links0.SetMultipleAttributes(atts, Links1.GetMultipleAttributes(atts))

    #Open reverse links if UP, BIDIRECT or TRCODE >= 50
    Links0.GetFilteredSet('[is_REVERSE]|[is_BIDIRECT]').SetAllAttValues('TypeNo', 1)

    #Close original links if UP
    Links1.GetFilteredSet('[is_REVERSE]').SetAllAttValues('TypeNo', 0)

    #Set TSys for open links
    for i, row in TSysDefs.iterrows():
        myTSys = Visum.Net.AddTSystem(i, 'PUT')
        myTSys.SetAttValue('Name', row['Name'])

    # Add the PuT-Aux transport system
    Tsys = Visum.Net.AddTSystem('PuTAux', 'PUTAUX')
    Tsys.SetAttValue("Name", "PuTAux")

    Visum.Net.Modes.ItemByKey('X').SetAttValue('TSysSet', PTpermitted)
    Visum.Net.Links.GetFilteredSet('[is_CLOSED]').SetAllAttValues('TypeNo', 0)
    Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetAllAttValues('TSysSet', railBased)
    Visum.Net.Turns.GetFilteredSet(f'[FromLink\\TSysSet]="{railBased}"&[ToLink\\TSysSet]="{railBased}"').SetAllAttValues('TSysSet', railBased)

def overpass_query(overpassQLstring):

    '''
    PLACEHOLDER.
        
        Parameters:
            overpassQLstring (string): PLACEHOLDER
    
        Returns:
            result (): PLACEHOLDER 
    '''
    
    #Create API object and boolean switch
    apiPy = overpy.Overpass()
    unsatis = True
    i=0
    
    #Keep trying to access the API until successful
    while unsatis:
        try:
            result = apiPy.query(overpassQLstring)
            unsatis = False
        
        #N.B. This except is generic, so will go into an infinite loop if internet connection is down or if overpassQLstring is invalid format
        except:
            time.sleep(2)
            i+=1
            if i < 10:
                print(traceback.format_exc())
                pass
            else:
                exit(0)
    
    #Return the API query result
    return result

def str_clean(myStr, desc):

    '''
    Takes a raw platform string from the OSM query data and cleans it to prepare for attempted match with BPLAN PLTs.
        
        Parameters:
            myStr (string): Contains an uncleaned OSM platfom string
            desc (string): Contains the station name
    
        Returns:
            myStr (string): A cleaned OSM platform string
    '''
    
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

def process_platformWays(myPlatformWays, crs, desc, x, y):
    
    '''
    Processes the platform data from OSM within a 500m x 500m bounding box centred on a given TIPLOC location
        
        Parameters:
            myPlatformWays (): PLACEHOLDER
            crs (string): PLACEHOLDER
            desc (string): PLACEHOLDER
            x (float): the X-Coordinate of the given TIPLOC
            y (float): the Y-Coordinate of the given TIPLOC
    
        Returns:
            myPlatformWays (): PLACEHOLDER 
            fig (pyplot object): A figure showing the processed platforms from OSM
    '''
    
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
            way.shape = LineString(way.bngs)
        else:
            way.shape = Polygon(way.bngs)
        
        #Cast as a geopandas object and add to our pre-defined plot and get the minimum rotated bounding rectangle
        gds = gpd.GeoSeries(way.shape)
        gds.plot(edgecolor = 'black', color = 'lightcoral', ax = ax2)
        mrr = way.shape.minimum_rotated_rectangle

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
                way.bng = nearest_points(way.shape, mrr_bisect.centroid)[0]
            else:
                way.bng = gpd.clip(gpd.GeoSeries(mrr_bisect), gds).centroid.values[0]
        
        #Calculate the distance from the OSM node and plot the platform centroid locations
        way.dist = way.bng.distance(Point(x, y))
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

def get_OSM_platform_data(path, TIPLOC, desc, x, y, bound):
    
    min = OSGB36toWGS84(x - bound, y - bound)
    max = OSGB36toWGS84(x + bound, y + bound)
    platformWays = overpass_query(f'way({min[0]},{min[1]},{max[0]},{max[1]})["railway"~"platform"];(._;>;);out body;').ways
    platformRelations = overpass_query(f'relation({min[0]},{min[1]},{max[0]},{max[1]})["railway"~"platform"];(._;>;);out body;').relations
    platformWays = platformWays + platformRelations

    #Process the OSM platform way data and save the resultant figure as a png file
    platformWays, myFig = process_platformWays(platformWays, TIPLOC, desc, x, y)
    myFig.savefig(os.path.join(path, f'cached_data\\OSM\\images\\{TIPLOC}.png'))
    
    #Convert the processed OSM platform way data into a pandas DataFrame and return alongside OSM station node location
    c1 = [way.tags['ref'] for way in platformWays]
    c2 = [way.bng for way in platformWays]
    c3 = [way.dist for way in platformWays]
    c4 = [way.shape for way in platformWays]
    myCols = {'Platform': c1, 'Location': c2, 'Dist': c3,  'Shape': c4}
    myTypes = {'Platform': str}
    dfPlatforms = pd.DataFrame.from_dict(myCols).astype(myTypes)
    dfPlatforms = dfPlatforms.explode('Platform').reset_index(drop = True)
    dfPlatforms = dfPlatforms.join(dfPlatforms.pop('Platform').str.split(';', expand = True))
    dfPlatforms = dfPlatforms.melt(dfPlatforms.columns[:len(myCols)-1], dfPlatforms.columns[len(myCols)-1:])
    dfPlatforms = dfPlatforms.rename(columns = {'value': 'Platform'}).drop('variable', axis = 1).sort_values('Dist').reset_index(drop = True)

    return dfPlatforms

def addStopPoint(Visum, i, row, bound, TPEsUnique):
    
    aTPE = TPEsUnique.loc[row['index_TIPLOC']]

    #Define a boolean that will indicate that the stop point has not yet been successfully added
    unsatis = True

    #Define an array for shifting linear interpolation to ensure that a stop point never has a RelPos witin 0.001 of the end of a link, the centrepoint of a link, or another stop point
    alt = [0, 0, 0, 0]

    #Create a visum map match object and find the nearest active link to the stop point location
    MyMapMatcher = Visum.Net.CreateMapMatcher()
    sp_Link = MyMapMatcher.GetNearestLink(row['Easting'], row['Northing'], bound, True, True)
    
    #If a match is found, determine whether the link should be directed or not. Otherwise match again without using attribute filter.
    try:
        is_dir = sp_Link.Link.AttValue('ReverseLink\\TypeNo') == 0
    except:
        Visum.Log(12288, f"No link within {bound}m for {row['index_TIPLOC']}: {aTPE['Tiploc']}: {aTPE['Name']} - Platform {row['PlatformID']}.")

    #Define the access node depending on the Relative Position calculated
    if sp_Link.RelPos < 0.5:
        sa_Node = sp_Link.Link.AttValue('FromNodeNo')
    else:
        sa_Node = sp_Link.Link.AttValue('ToNodeNo')

    #! above does not take account of links that are then split and therefore SP node number != SA node number. When nodes are split, SA numbers should be updated to reflect this

    #Add a new stop area for the platform and populate attributes
    sa = Visum.Net.AddStopArea(i, row['index_TIPLOC'], sa_Node, sp_Link.XPosOnLink, sp_Link.YPosOnLink)
    sa.SetAttValue('Code', f"{aTPE['Tiploc']}_{row['PlatformID']}")
    sa.SetAttValue('Name', f"Platform {row['PlatformID']}")
    
    #Attempt to add the stop point on the link, taking into account whether it should be directed or not
    try:
        sp = Visum.Net.AddStopPointOnLink(i, sa, sp_Link.Link.AttValue('FromNodeNo'), sp_Link.Link.AttValue('ToNodeNo'), is_dir)
    
    #If this fails, the link is split at the midpoint before attempting again
    except:
        sp_Node = Visum.Net.AddNode(i, sp_Link.Link.GetXCoordAtRelPos(0.5), sp_Link.Link.GetYCoordAtRelPos(0.5))
        Visum.Net.StopAreas.ItemByKey(i).SetAttValue('NodeNo', i)
        sp_Link.Link.SplitViaNode(sp_Node)
        sp_Link = MyMapMatcher.GetNearestLink(row['Easting'], row['Northing'], bound, True, True)
        sp = Visum.Net.AddStopPointOnLink(i, sa, sp_Link.Link.AttValue('FromNodeNo'), sp_Link.Link.AttValue('ToNodeNo'), is_dir)
    
    #After the stop point has been added, the RelPos is shifted according to areas of potential conflict to ensure other stop points can be added to the same link
    while unsatis:
        RelPos = interp1d([0, 0.5, 0.5, 1],[0 + alt[0], 0.5 - alt[1], 0.5 + alt[2], 1 - alt[3]])
        NewRelPos = float(RelPos(sp_Link.RelPos))
        shiftBool = [NewRelPos < 0.001, (NewRelPos > 0.497) & (NewRelPos <= 0.500), (NewRelPos >= 0.500) & (NewRelPos < 0.503), NewRelPos > 0.999]
        if np.any(shiftBool):
            alt = [altN + 0.001 if boolN else altN for altN, boolN in zip(alt, shiftBool)]
        else:
            try:
                sp.SetAttValue('RelPos', NewRelPos)
                unsatis = False
            except:
                alt = [altN + 0.001 for altN in alt]

    #The attributes for the stop point are populated
    sp.SetAttValue('Code', f"{aTPE['Tiploc']}_{row['PlatformID']}")
    sp.SetAttValue('Name', f"Platform {row['PlatformID']}")

def getJoin(x):
    xFil = x[x.notna()].unique()
    xList = [str(anX) for anX in xFil.tolist() if str(anX) != '<NA>']
    if len(xList) == 1:
        return xList[0]
    else:
        return ','.join(xList)
    
def getCommonPrefix(x):
    xFil = x[x.notna()].unique()
    xList = [str(anX) for anX in xFil.tolist() if str(anX) != '<NA>']
    if len(xList) == 1:
        return xList[0]
    else:
        return f'{os.path.commonprefix(xList)}*'

def processBPLAN(path, bplan_file, tiploc_file):

    myTPE = pd.DataFrame(json.load(open(tiploc_file)).get('Tiplocs')).set_index('Tiploc')
    myTPE = myTPE[[len(aTiploc) <= 7 for aTiploc in myTPE.index.values]]
    myTPE['CRS'] = [row['Details'].get('CRS') for _, row in myTPE.iterrows()]
    myTPE.drop(['DisplayName', 'NodeId', 'Codes', 'Details', 'Elevation'], axis = 1, inplace = True)
    myTPE['Easting'] = [WGS84toOSGB36(float(row['Latitude']), float(row['Longitude']))[0] for _, row in myTPE.iterrows()]
    myTPE['Northing'] = [WGS84toOSGB36(float(row['Latitude']), float(row['Longitude']))[1] for _, row in myTPE.iterrows()]
    myTPE['Stanox'] = myTPE['Stanox'].astype('Int32', False).astype('str', False)

    TPEsUnique = pd.pivot_table(myTPE.reset_index(), ['Tiploc', 'Name', 'Stanox', 'Latitude', 'Longitude', 'InBPlan', 'InTPS', 'CRS'], ['Easting', 'Northing'],
                               aggfunc = {'Tiploc': getJoin,
                                          'Name': getCommonPrefix,
                                          'Stanox': getJoin,
                                          'Latitude': np.mean,
                                          'Longitude': np.mean,
                                          'InBPlan': np.mean,
                                          'InTPS': np.mean,
                                          'CRS': getJoin}).reset_index().reset_index()

    TPEsUnique['index'] += 100000
    TPEsUnique.set_index(['index'], inplace = True)

    myTPE = myTPE.reset_index().merge(TPEsUnique.reset_index(), 'left', ['Easting', 'Northing']).set_index('Tiploc_x')

    myTPE.to_csv(os.path.join(path, 'cached_data\\BPLAN\\TPEs.csv'))
    
    with open(bplan_file) as f:
        lines = f.readlines()
    lines = [line[:-1].split('\t') for line in lines]

    PLTs = pd.DataFrame(list(filter(lambda line: line[0] == "PLT", lines)),
                        columns = ['RecordType', 'ActionCode', 'TIPLOC', 'PlatformID', 'StartDate', 'EndDate',
                                   'PlatformLength', 'PowerSupplyType', 'PassengerDOO', 'NonPassengerDOO'])
    
    PLTs.drop(['RecordType', 'ActionCode', 'EndDate', 'PowerSupplyType'], axis = 1, inplace = True)
    PLTs = PLTs[[aTIPLOC in myTPE.index for aTIPLOC in PLTs['TIPLOC']]]
    PLTs['index_TIPLOC'] = myTPE.loc[PLTs['TIPLOC']]['index'].values
    
    PLTs = PLTs[~PLTs['index_TIPLOC'].isna()]
    PLTs['PlatformID'] = PLTs['PlatformID'].str.upper()
    PLTs['index_PlatformID'] = [sorted(PLTs['PlatformID'].unique()).index(PlatformID) + 1 for PlatformID in PLTs['PlatformID']]
    PLTs['index'] = 1000*PLTs['index_TIPLOC']
    PLTs['TIPLOC_PlatformID'] = PLTs['TIPLOC'] + '_' +  PLTs['PlatformID']
    PLTs.set_index(['TIPLOC_PlatformID'], inplace = True)
    PLTs['StartDate'] = pd.to_datetime(PLTs['StartDate'], dayfirst = True)
    PLTs['PlatformLength'] = PLTs['PlatformLength'].replace('', 0)
    PLTs['PlatformLength'] = PLTs['PlatformLength'].astype('int32', False)
    
    for col in ['PassengerDOO', 'NonPassengerDOO']:
        PLTs[col] = [aBool == "Y" for aBool in PLTs[col]]
    
    for col in ['Easting', 'Northing', 'Quality']:
        PLTs[col] = 0

    ex = wx.App()

    prog = wx.ProgressDialog("Platforms", "Getting OSM platform locations...",
                                            len(PLTs),
                                            style=wx.PD_APP_MODAL | wx.PD_SMOOTH | wx.PD_AUTO_HIDE)
    for j, (i, row) in enumerate(PLTs.iterrows()):
        myPlatformNum = re.sub('[^0-9]', '', row['PlatformID'])
        aTPE = myTPE.loc[row['TIPLOC']]
        OSMpltData = get_OSM_platform_data(path, row['index_TIPLOC'], f"{aTPE['Tiploc_y']}: {aTPE['Name_y']}", aTPE['Easting'], aTPE['Northing'], 250)
        OSMpltDataFil = OSMpltData[OSMpltData['Platform'] == row['PlatformID']]
        if len(OSMpltDataFil) > 0:
            OSMloc = OSMpltDataFil['Location'].iloc[0]
            PLTs.loc[i, 'Easting'] = OSMloc.x
            PLTs.loc[i, 'Northing'] = OSMloc.y
            PLTs.loc[i, 'Shape'] = OSMpltDataFil['Shape'].iloc[0]
            PLTs.loc[i, 'Quality'] = 2
            PLTs.loc[i, 'index'] += row['index_PlatformID']
        elif len(myPlatformNum) > 0:
            OSMpltData['Platform'] = [re.sub('[^0-9]', '', str(platform)) for platform in OSMpltData['Platform']]
            OSMpltDataFil = OSMpltData[OSMpltData['Platform'] == myPlatformNum]
            if len(OSMpltDataFil) > 0:
                OSMloc = OSMpltDataFil['Location'].iloc[0]
                PLTs.loc[i, 'Easting'] = OSMloc.x
                PLTs.loc[i, 'Northing'] = OSMloc.y
                PLTs.loc[i, 'Shape'] = OSMpltDataFil['Shape'].iloc[0]
                PLTs.loc[i, 'Quality'] = 1
                PLTs.loc[i, 'index'] += row['index_PlatformID']
        prog.Update(j, f"Getting OSM platform locations... ({int(j)}/{len(PLTs)})")
    prog.Destroy()
    
    PLTs = PLTs[PLTs['Quality'] > 0]
    PLTs.to_csv(os.path.join(path, 'cached_data\\BPLAN\\PLTs.csv'))
    
    PLTs= pd.read_csv(os.path.join(path, 'cached_data\\BPLAN\\PLTs.csv'),index_col=0)
    PLTsUnique = pd.pivot_table(PLTs.reset_index(), ['PlatformID', 'StartDate', 'PlatformLength', 'PassengerDOO', 'NonPassengerDOO',
                                                     'index_TIPLOC', 'index_PlatformID', 'Easting', 'Northing', 'Quality'],
                                ['index'],
                               aggfunc = {'PlatformID': getCommonPrefix,
                                          'StartDate': np.min,
                                          'PlatformLength': np.min,
                                          'PassengerDOO': np.mean,
                                          'NonPassengerDOO': np.mean,
                                          'index_TIPLOC': np.min,
                                          'index_PlatformID': np.min,
                                          'Easting': np.min,
                                          'Northing': np.min,
                                          'Quality': np.min})

    return TPEsUnique, PLTsUnique

def getVisumLOCs(path, TPEsUnique, myVer, myShp, reversedELRs, tsys_path):
    Visum = com.Dispatch('Visum.Visum.240')
    Visum.SetPath(57, os.path.join(path,f"cached_data"))
    Visum.SetLogFileName(f"Log_LOCs_{datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')}.txt")
    projString = """
                        PROJCS[
                            "British_National_Grid_TOWGS",
                            GEOGCS[
                                "GCS_OSGB_1936",
                                DATUM[
                                    "D_OSGB_1936",
                                    SPHEROID["Airy_1830",6377563.396,299.3249646],
                                    TOWGS84[446.4,-125.2,542.1,0.15,0.247,0.842,-20.49]
                                ],
                                PRIMEM["Greenwich",0],
                                UNIT["Degree",0.017453292519943295]
                            ],
                            PROJECTION["Transverse_Mercator"],
                            PARAMETER["False_Easting",400000],
                            PARAMETER["False_Northing",-100000],
                            PARAMETER["Central_Meridian",-2],
                            PARAMETER["Scale_Factor",0.999601272],
                            PARAMETER["Latitude_Of_Origin",49],
                            UNIT["Meter",1]
                        ]
                    """
    try:
        Visum.Net.SetProjection(projString, False)
        Visum.Net.SetAttValue('LeftHandTraffic', 1)
        ImportShapeFilePara = Visum.IO.CreateImportShapeFilePara()
        ImportShapeFilePara.CreateUserDefinedAttributes = True
        ImportShapeFilePara.ObjectType = 0
        ImportShapeFilePara.SetAttValue('Directed', True)
        Visum.IO.ImportShapefile(myShp, ImportShapeFilePara)
        TSysDefs = pd.read_csv(tsys_path, low_memory = False).set_index('Code')
        railBased = str(TSysDefs.index[TSysDefs['rail_based']].values).replace('\n', '').replace("' '", ',').replace("['", '').replace("']", '')
        PTpermitted = str(TSysDefs.index[TSysDefs['PT_permitted']].values).replace('\n', '').replace("' '", ',').replace("['", '').replace("']", '') + ',W,PuTAux' 
        fixDirectedNet(Visum, reversedELRs, TSysDefs, railBased, PTpermitted)
        MyMapMatcher = Visum.Net.CreateMapMatcher()
        LinkType = Visum.Net.AddLinkType(2)
        LinkType.SetAttValue('TSysSet', railBased)
        for uda, dtype in [['CRS', 5],['InBPlan', 2],['InTPS', 2],['Stanox', 5]]:
            Visum.Net.Stops.AddUserDefinedAttribute(uda, uda, uda, dtype)
        Visum.Graphic.StopDrawing = True
        ex = wx.App()
        TPEoffset = TPEsUnique.index.min()


        dfNodes = TPEsUnique[['Easting', 'Northing', 'Tiploc', 'Name']].copy()
        dfNodes.reset_index(inplace=True)
        dfNodes.rename({'index':'$NODE:NO', 'Easting':'XCOORD', 'Northing':'YCOORD', 'Tiploc':'Code'}, axis=1, inplace=True)
        
        dfStops = TPEsUnique[['Easting', 'Northing', 'Tiploc', 'Name', 'CRS', 'InBPlan', 'InTPS', 'Stanox']].copy()
        dfStops.reset_index(inplace=True)
        dfStops.rename({'index':'$STOP:NO', 'Easting':'XCOORD', 'Northing':'YCOORD', 'Tiploc':'Code'}, axis=1, inplace=True)

        dfStopAreasPU = TPEsUnique[['Easting', 'Northing']].copy()
        dfStopAreasPU.reset_index(inplace=True)
        dfStopAreasPU['NodeNo'] = dfStopAreasPU['index']
        dfStopAreasPU['$STOPAREA:NO'] = dfStopAreasPU['index']*1000
        dfStopAreasPU['Name'] = 'Platform Unknown'
        dfStopAreasPU.rename({'index':'StopNo', 'Easting':'XCOORD', 'Northing':'YCOORD'}, axis=1, inplace=True)
        dfStopAreasPU = dfStopAreasPU[['$STOPAREA:NO', 'NodeNo', 'StopNo', 'XCOORD', 'YCOORD', 'Name']]

        dfStopAreasAE = TPEsUnique[['Easting', 'Northing']].copy()
        dfStopAreasAE.reset_index(inplace=True)
        dfStopAreasAE['NodeNo'] = dfStopAreasAE['index']
        dfStopAreasAE['$STOPAREA:NO'] = dfStopAreasAE['index']*1000+999
        dfStopAreasAE['Name'] = 'AccessEgress'
        dfStopAreasAE.rename({'index':'StopNo', 'Easting':'XCOORD', 'Northing':'YCOORD'}, axis=1, inplace=True)
        dfStopAreasAE = dfStopAreasAE[['$STOPAREA:NO', 'NodeNo', 'StopNo', 'XCOORD', 'YCOORD', 'Name']]

        dfStopAreas = pd.concat([dfStopAreasPU, dfStopAreasAE])

        dfStopPoints = TPEsUnique[['Tiploc']].copy()
        dfStopPoints.reset_index(inplace=True)
        dfStopPoints['$STOPPOINT:NO'] = dfStopPoints['index']*1000
        dfStopPoints['StopAreaNo'] = dfStopPoints['index']*1000
        dfStopPoints['Name'] = 'Platform Unknown'
        dfStopPoints.rename({'index':'NodeNo', 'Tiploc':'Code'}, axis=1, inplace=True)
        dfStopPoints = dfStopPoints[['$STOPPOINT:NO', 'StopAreaNo', 'NodeNo', 'Name', 'Code']]
        
        # write .net file
        header = '''$VISION

$VERSION:VERSNR;FILETYPE;LANGUAGE;UNIT
15;Net;ENG;KM

'''

        with open(os.path.join(path, 'cached_data\\VISUM\\LOCS.net'), 'w') as f:
            f.write(header)
            dfNodes.to_csv(f, mode='a', sep=';', index=False, lineterminator='\n')
            f.write('\n')
            dfStops.to_csv(f, mode='a', sep=';', index=False, lineterminator='\n')
            f.write('\n')
            dfStopAreas.to_csv(f, mode='a', sep=';', index=False, lineterminator='\n')
            f.write('\n')
            dfStopPoints.to_csv(f, mode='a', sep=';', index=False, lineterminator='\n')

        Visum.IO.SaveVersion(myVer)
        Visum.IO.LoadNet(os.path.join(path,'cached_data\\VISUM\\LOCS.net'), True)


        prog = wx.ProgressDialog("TIPLOCs", "Generating TIPLOC objects...",
                                            TPEsUnique.index.max() - TPEoffset,
                                            style=wx.PD_APP_MODAL | wx.PD_SMOOTH | wx.PD_AUTO_HIDE)

        for i, row in TPEsUnique.iterrows():
            unsatis = True
            fil_string = '[TYPENO]=1'
            nTRID = 0
            while unsatis & (nTRID < 10):
                Visum.Net.Links.SetPassive()
                Visum.Net.Links.GetFilteredSet(fil_string).SetActive()
                split_Link = MyMapMatcher.GetNearestLink(row['Easting'], row['Northing'], 250, True, True)
                unsatis = split_Link.Success
                if unsatis:
                    split_TRID = split_Link.Link.AttValue('TRID')
                    split_no = 10*i + nTRID
                    if split_Link.RelPos == 0:
                        try:
                            Visum.Net.AddLink(split_no, split_Link.Link.AttValue('FromNodeNo'), i, 2)
                        except:
                            pass
                    elif split_Link.RelPos == 1:
                        try:
                            Visum.Net.AddLink(split_no, split_Link.Link.AttValue('ToNodeNo'), i, 2)
                        except:
                            pass
                    else:
                        split_Node = Visum.Net.AddNode(split_no, split_Link.XPosOnLink, split_Link.YPosOnLink)
                        split_Link.Link.SplitViaNode(split_Node)
                        Visum.Net.Links.ItemByKey(split_Link.Link.AttValue('FromNodeNo'), split_no).SetNo(split_no + 10*TPEoffset)
                        Visum.Net.Links.ItemByKey(split_no, split_Link.Link.AttValue('ToNodeNo')).SetNo(split_no + 20*TPEoffset)
                        Visum.Net.AddLink(split_no, split_no, i, 2)
                    fil_string += f"&[TRID]!=\"{split_TRID}\""
                    nTRID += 1
            prog.Update(i- TPEoffset, f"Generating TIPLOC objects... ({int(i- TPEoffset)}/{int(TPEsUnique.index.max() - TPEoffset)})")
        prog.Destroy()
        Visum.Graphic.StopDrawing = False
        Visum.Net.Turns.GetFilteredSet('[FromLink\\TypeNo]=2&[ToLink\\TypeNo]=2&[FromLink\\No]!=[ToLink\\No]').SetAllAttValues('TSysSet', '')
        Visum.IO.SaveVersion(myVer)
    except:
        Visum.Log(12288, traceback.format_exc())



def getVisumPLTs(PLTsUnique, myPLTsVer, myLOCsVer, TPEsUnique, output):
    Visum = com.Dispatch('Visum.Visum.240')
    Visum.SetPath(57, os.path.join(path,f"cached_data"))
    Visum.SetLogFileName(f"Log_PLTs_{datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')}.txt")
    try:
        Visum.IO.LoadVersion(myLOCsVer)
        Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetActive()
        Visum.Graphic.StopDrawing = True
        ex = wx.App()
        PLToffset = PLTsUnique.index.min()
        prog = wx.ProgressDialog("Platforms", "Generating platform objects...",
                                            PLTsUnique.index.max() - PLToffset,
                                            style=wx.PD_APP_MODAL | wx.PD_SMOOTH | wx.PD_AUTO_HIDE)
        for i, row in PLTsUnique.iterrows():
            addStopPoint(Visum, i, row, 250, TPEsUnique)
            prog.Update(i- PLToffset, f"Generating platform objects... ({int(i- PLToffset)}/{int(PLTsUnique.index.max() - PLToffset)})")
        prog.Destroy()
        Visum.Graphic.StopDrawing = False
        Visum.Net.Links.SetMultipleAttributes(['Length'], Visum.Net.Links.GetMultipleAttributes(['LengthPoly']))
        DFcols_Visum = ['No', 'Code', 'Name', 'YCoord', 'XCoord']
        DFcols_GTFS = ['stop_id', 'stop_code', 'stop_name', 'stop_lat', 'stop_lon']
        DF_s = pd.DataFrame(Visum.Net.Stops.GetMultipleAttributes(DFcols_Visum), columns = DFcols_GTFS)
        DF_s['location_type'] = 1
        DF_sp = pd.DataFrame(Visum.Net.StopPoints.GetMultipleAttributes(DFcols_Visum + ['StopArea\\StopNo']), columns = DFcols_GTFS + ['parent_station'])
        DF_sp['location_type'] = 0
        DF_full = pd.concat([DF_s, DF_sp], axis = 0)
        DF_full['stop_id'] = pd.to_numeric(DF_full['stop_id'],errors='coerce').astype('Int64')
        DF_full['parent_station'] = pd.to_numeric(DF_full['parent_station'], errors = 'coerce').astype('Int64')
        DF_full.to_csv(output, index = False)
        Visum.IO.SaveVersion(myPLTsVer)
    except:
        Visum.Log(12288, traceback.format_exc())

def addZonesandConnectors(Visum):
    # Add zones on top of platform unknown locations for stops where a CRS code is defined and add a connector between this zone and the Platform Unknown 

    allStopAreas = Visum.Net.StopAreas.FilteredBy(f'[NAME]="Platform Unknown"&[STOP\CRS]!=""&[STOP\CODE]=[CODE]')
    atts = ['Code', 'NodeNo', 'XCoord', 'YCoord', 'Stop\\Name', 'Stop\\CRS']
    allStopAreasDF = pd.DataFrame(allStopAreas.GetMultipleAttributes(atts), columns = atts).set_index('Code')

    #Turn off Visum drawing to imporve performance
    Visum.Graphic.StopDrawing = True

    #Now start iterating through Zones
    for i, row in allStopAreasDF.iterrows():
        
        #Add the CRS zone and define code & name, before providing the connector
        aZone = Visum.Net.AddZone(-1, row['XCoord'], row['YCoord'])
        aZone.SetAttValue('Code', row['Stop\\CRS'])
        aZone.SetAttValue('Name', row['Stop\\Name'])
        Visum.Net.AddConnector(aZone, row['NodeNo'])

    return allStopAreasDF

def addTransferLinks(Visum, xfer_link_path, allStopAreasDF):
    #Add a new link type for use with walk transfer links
    LinkType = Visum.Net.AddLinkType(3)
    LinkType.SetAttValue('TSysSet', 'W')
    LinkType.SetAttValue('Name', 'WalkTransferLink')

    #Add another link type to be used for automatic walking speed transfer links
    LinkType = Visum.Net.AddLinkType(4)
    LinkType.SetAttValue('TSysSet', 'W')
    LinkType.SetAttValue('Name', 'WalkProximityLink')


    # Add a new link type for use with PuT-Aux transfer links
    LinkType = Visum.Net.AddLinkType(5)
    LinkType.SetAttValue("TSysSet", "PuTAux")
    LinkType.SetAttValue("Name", "PuTAuxTransferLink")
    
    #Make a DataFrame of user defined transfer links and set the index
    myCSV = pd.read_csv(xfer_link_path, low_memory = False).set_index(['FromCRS', 'ToCRS'])
    
    #Iterate through the user defined transfer links
    for i, row in myCSV.iterrows():

        #We only need to create one if the from_CRS alphabetically preceeds the to_CRS, because the reverse direction will automatically be created
        if i[0] < i[1]:

            try:
                myFromNode = allStopAreasDF[allStopAreasDF['Stop\\CRS'] == i[0]]['NodeNo'][0]
                fromFlag = True
            except:
                Visum.Log(16384, f'No served node found for {i[0]}. No transfer link will be created unless you define the desired CRS-TIPLOC match in the manual override csv.')
                fromFlag = False
        
            try:
                myToNode = allStopAreasDF[allStopAreasDF['Stop\\CRS'] == i[1]]['NodeNo'][0]
                toFlag = True
            except:
                Visum.Log(16384, f'No served node found for {i[1]}. No transfer link will be created unless you define the desired CRS-TIPLOC match in the manual override csv.')
                toFlag = False

            #If both from_CRS and to_CRS are found, create the user defined transfer link and apply the correct times to both directions
            if fromFlag & toFlag:
                if row.TSys == "Walk":
                    myLink = Visum.Net.AddLink(-1, myFromNode, myToNode, 3)
                    myLink.SetAttValue('T_PUTSYS(W)', 60*row['TravelTime'])
                    try:
                        myLink.SetAttValue('REVERSELINK\\T_PUTSYS(W)', 60*myCSV.loc[(i[1], i[0]), 'TravelTime'])
                    except:
                        myLink.SetAttValue('REVERSELINK\\T_PUTSYS(W)', 60*myCSV.loc[(i[0], i[1]), 'TravelTime'])
                elif row.TSys == "PuT-Aux":
                    myLink = Visum.Net.AddLink(-1, myFromNode, myToNode, 5)
                    myLink.SetAttValue('T_PUTSYS(PuTAux)', 60*row['TravelTime'])
                    try:
                        myLink.SetAttValue('REVERSELINK\\T_PUTSYS(PuTAux)', 60*myCSV.loc[(i[1], i[0]), 'TravelTime'])
                    except:
                        myLink.SetAttValue('REVERSELINK\\T_PUTSYS(PuTAux)', 60*myCSV.loc[(i[0], i[1]), 'TravelTime'])
                else:
                    Visum.Log(16384, f'Tsys {row.TSys} does not exist in the network. No transfer link between {i[0]} and {i[1]} created.')


    #Create a list of all possible cobinations of served dummy Stop Areas
    cc = list(combinations(allStopAreasDF[['NodeNo', 'XCoord', 'YCoord']].values, 2))

    #Iterate through each possible permutation and calculate the distance, adding a link for any pair closer than 250m (N.B. This is many orders of magnitude faster then the Visum COM MapMatcher geographic search approach)
    for pair in cc:
        nodeFrom, xFrom, yFrom = pair[0]
        nodeTo, xTo, yTo = pair[1]
        distance = math.sqrt((xTo - xFrom)**2 + (yTo - yFrom)**2)
        if (distance < 250) & (nodeFrom < nodeTo):
            try:
                Visum.Net.AddLink(-1, nodeFrom, nodeTo, 4)
            except:
                Visum.Log(16384, "This link has already been been manually defined. Therefore, no link is created.")
    
    #We have now finished iterative slow processes, so we can turn back on drawing in Visum again
    Visum.Graphic.StopDrawing = False


def update_crs(Visum, crs_path):
    df_crs = pd.read_csv(crs_path)

    stops = pd.DataFrame(Visum.Net.Stops.GetMultipleAttributes(['NO', 'CODE', 'CRS']), columns=['No', 'Code', 'CRS'])

    for i, row in df_crs.iterrows():
        tiploc = row['StopCode']
        stopno = stops.loc[stops.Code == tiploc].No.tolist()[0]
        stop = Visum.Net.Stops.ItemByKey(stopno)

        if row.NewCoordinates == 1:
            stop.SetAttValue('CRS', row.NewCRS)
        else:
            stop.SetAttValue('CRS', "")



def main(path, myShp, tiploc_path, BPLAN_path, ELR_path, merge_path, tsys_path, xfer_link_path):
    
    # Reprocess BPLAN to obtain new pickle results and save to cache
    TPEsUnique, PLTsUnique = processBPLAN(path, BPLAN_path, tiploc_path)

    myPickle = os.path.join(path, 'cached_data\\BPLAN\\uniques.p')
    with open(myPickle, 'wb') as f:
        pickle.dump([TPEsUnique, PLTsUnique], f)
    
    
    myLOCsVer = os.path.join(path, 'cached_data\\VISUM\\LOCs_Only.ver')
    myPLTsVer = os.path.join(path, 'cached_data\\VISUM\\LOCs_and_PLTs.ver')
    output = os.path.join(path, 'output\\GTFS\\stops.txt')
    
    # Reprocess BPLAN to obtain new LOCs Version file and save to cache
    reversedELRsDF = pd.read_csv(ELR_path, low_memory = False)
    reversedELRs = [reversedELR[0] for reversedELR in reversedELRsDF.values]
    getVisumLOCs(path, TPEsUnique, myLOCsVer, myShp, reversedELRs, tsys_path)


    # Reprocess BPLAN to obtain new PLTs Version file and save to cache
    getVisumPLTs(PLTsUnique, myPLTsVer, myLOCsVer, TPEsUnique, output)
    
    
    Visum = com.Dispatch("Visum.Visum.240")
    Visum.SetPath(57, os.path.join(path,f"cached_data"))
    Visum.SetLogFileName(f"Log_LOCs_and_PLTs_{datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')}.txt")
    try:
        Visum.LoadVersion(myPLTsVer)
        # Update CRS codes from override file
        update_crs(Visum, merge_path)

        # Create zones and connectors for CRS stops
        allStopAreasDF = addZonesandConnectors(Visum)

        # Add transfer links to the network
        addTransferLinks(Visum, xfer_link_path, allStopAreasDF)

        Visum.Net.SetAttValue("STRONGLINEROUTELENGTHSADAPTION", 1)

        Visum.SaveVersion(os.path.join(path, 'cached_data\\VISUM\\LOCs_and_PLTs_ZonesConnectorsXferLinks.ver'))
    except:
        Visum.Log(12288, traceback.format_exc())



if __name__ == "__main__":

    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    buildNetwork = gi.readNetworkInputs(input_path)
    if bool(buildNetwork[0]):
        myShp = buildNetwork[1]
        tiploc_path = buildNetwork[2]
        BPLAN_path =  buildNetwork[3]
        ELR_path = buildNetwork[4]
        merge_path = buildNetwork[5]
        tsys_path = buildNetwork[6]
        xfer_link_path = buildNetwork[7]

        main(path, myShp, tiploc_path, BPLAN_path, ELR_path, merge_path, tsys_path, xfer_link_path)
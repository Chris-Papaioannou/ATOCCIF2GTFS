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

def fixDirectedNet(Visum):

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
    Visum.Net.AddTSystem('2', 'PUT')
    Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetAllAttValues('TSysSet', '2')
    Visum.Net.Turns.GetFilteredSet('[FromLink\\TSysSet]="2"&[ToLink\\TSysSet]="2"').SetAllAttValues('TSysSet', '2')

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

def process_platformWays(myPlatformWays, crs, desc, x, y):
    
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
    
    #Check if OSM data has been saved as pickle already for this station, and if so, read from pickle file, otherwise query them from OSM
    myPickle = os.path.join(path, f'OSM_pickles\\{TIPLOC}_platforms.p')
    if os.path.exists(myPickle):
        with open(myPickle, 'rb') as f:
            dfPlatforms, = pickle.load(f)
    else:
        min = OSGB36toWGS84(x - bound, y - bound)
        max = OSGB36toWGS84(x + bound, y + bound)
        platformWays = overpass_query(f'way({min[0]},{min[1]},{max[0]},{max[1]})["railway"~"platform"];(._;>;);out body;').ways
        platformRelations = overpass_query(f'relation({min[0]},{min[1]},{max[0]},{max[1]})["railway"~"platform"];(._;>;);out body;').relations
        platformWays = platformWays + platformRelations
    
        #Process the OSM platform way data and save the resultant figure as a png file
        platformWays, myFig = process_platformWays(platformWays, TIPLOC, desc, x, y)
        myFig.savefig(os.path.join(path, f'OSM_images\\{TIPLOC}_platforms.png'))
        
        #Convert the processed OSM platform way data into a pandas DataFrame and return alongside OSM station node location
        c1 = [way.tags['ref'] for way in platformWays]
        c2 = [way.bng for way in platformWays]
        c3 = [way.dist for way in platformWays]
        myCols = {'Platform': c1, 'Location': c2, 'Dist': c3}
        myTypes = {'Platform': str}
        dfPlatforms = pd.DataFrame.from_dict(myCols).astype(myTypes)
        dfPlatforms = dfPlatforms.explode('Platform').reset_index(drop = True)
        dfPlatforms = dfPlatforms.join(dfPlatforms.pop('Platform').str.split(';', expand = True))
        dfPlatforms = dfPlatforms.melt(dfPlatforms.columns[:len(myCols)-1], dfPlatforms.columns[len(myCols)-1:])
        dfPlatforms = dfPlatforms.rename(columns = {'value': 'Platform'}).drop('variable', axis = 1).sort_values('Dist').reset_index(drop = True)
        with open(myPickle, 'wb') as f:
            pickle.dump([dfPlatforms], f)
    return dfPlatforms

def addStopPoint(Visum, i, row, bound, LOCsUnique):
    
    myLOC = LOCsUnique.loc[row['index_TIPLOC']]

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
        print(f"ERROR: No link within {bound}m for {row['index_TIPLOC']}: {myLOC['TIPLOC']}: {myLOC['LocationName']} - Platform {row['PlatformID']}.")

    #Define the access node depending on the Relative Position calculated
    if sp_Link.RelPos < 0.5:
        sa_Node = sp_Link.Link.AttValue('FromNodeNo')
    else:
        sa_Node = sp_Link.Link.AttValue('ToNodeNo')

    #Add a new stop area for the platform and populate attributes
    sa = Visum.Net.AddStopArea(i, row['index_TIPLOC'], sa_Node, sp_Link.XPosOnLink, sp_Link.YPosOnLink)
    sa.SetAttValue('Code', f"{myLOC['TIPLOC']}_{row['PlatformID']}")
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
    
    #While the stop point has not yet been added, the RelPos is shifted according to areas of potential conflict to ensure other stop points can be added to the same link
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

    #The attributes for the stop point are populated
    sp.SetAttValue('Code', f"{myLOC['TIPLOC']}_{row['PlatformID']}")
    sp.SetAttValue('Name', f"Platform {row['PlatformID']}")

def getCommonPrefix(x):
    xList = [str(anX) for anX in x.tolist()]
    if len(xList) == 1:
        return xList[0]
    else:
        return f'{os.path.commonprefix(xList)}*'

def processBPLAN(path):

    NaPTANstops = pd.read_csv(os.path.join(path, 'input\\Stops.csv'), index_col = 'ATCOCode', low_memory = False)
    NaPTANstops = NaPTANstops[NaPTANstops['StopType'] == 'RLY']

    with open(os.path.join(path, 'input\\Geography_20221210_to_20230520_from_20221211.txt')) as f:
        lines = f.readlines()
    lines = [line[:-1].split('\t') for line in lines]
    
    LOCs = pd.DataFrame(list(filter(lambda line: line[0] == "LOC", lines)),
                        columns = ['RecordType', 'ActionCode', 'TIPLOC', 'LocationName', 'StartDate', 'EndDate',
                                   'Easting', 'Northing', 'TimingPointType', 'ZoneResponsible', 'STANOX', 'OffNetwork', 'ForceLPB'])
    
    LOCs.drop(['RecordType', 'ActionCode', 'EndDate', 'ForceLPB'], axis = 1, inplace = True)
    LOCs.set_index('TIPLOC', inplace = True)
    LOCs['StartDate'] = LOCs['StartDate'].astype('datetime64', False)
    LOCs['STANOX'] = LOCs['STANOX'].replace('', 0)
    
    for col in ['Easting', 'Northing', 'ZoneResponsible', 'STANOX']:
        LOCs[col] = LOCs[col].astype('int32', False)
    LOCs['OffNetwork'] = [aBool == "Y" for aBool in LOCs['OffNetwork']]
    
    LOCs['Quality'] = 0
    
    for i, row in LOCs.iterrows():
        try:
            NaPTANcheck = NaPTANstops.loc['9100' + i]
            LOCs.loc[i, 'Easting'] = NaPTANcheck['Easting']
            LOCs.loc[i, 'Northing'] = NaPTANcheck['Northing']
            LOCs.loc[i, 'Quality'] = 4
        except KeyError:
            sharedSTANOX = LOCs[LOCs['STANOX'] == row['STANOX']].drop(['Easting', 'Northing'], axis = 1)
            if (row['STANOX'] > 0) & (len(sharedSTANOX) > 0):
                sharedSTANOX.index = '9100' + sharedSTANOX.index
                sharedSTANOX = sharedSTANOX.join(NaPTANstops, how = 'inner')
                if len(sharedSTANOX) > 0:
                    LOCs.loc[i, 'Easting'] = sharedSTANOX['Easting'].mean()
                    LOCs.loc[i, 'Northing'] = sharedSTANOX['Northing'].mean()
                    LOCs.loc[i, 'Quality'] = 3
    
    boundE = LOCs['Easting'].loc[['PENZNCE', 'LOWSTFT']]
    boundN = LOCs['Northing'].loc[['PENZNCE', 'THURSO']]
    untrustedCoords = [0, 1, 10000, 99999, 111111, 222222, 333333, 444444, 555555, 666666, 777777, 888888, 989898, 999999]
    
    LOCs.loc[LOCs['Easting'].between(boundE[0], boundE[1])
             & LOCs['Northing'].between(boundN[0], boundN[1])
             & ~LOCs['Easting'].isin(untrustedCoords)
             & ~LOCs['Northing'].isin(untrustedCoords)
             & (LOCs['Quality'] == 0), 'Quality'] = 2
    
    for i, row in LOCs.iterrows():
        if row['Quality'] == 0:
            sharedTrustedSTANOX = LOCs[(LOCs['STANOX'] == row['STANOX']) & (LOCs['Quality'] == 2)]
            if (row['STANOX'] > 0) & (len(sharedTrustedSTANOX) > 0):
                LOCs.loc[i, 'Easting'] = sharedTrustedSTANOX['Easting'].mean()
                LOCs.loc[i, 'Northing'] = sharedTrustedSTANOX['Northing'].mean()
                LOCs.loc[i, 'Quality'] = 1
    
    LOCsUnique = pd.pivot_table(LOCs[LOCs['Quality'] > 0].reset_index(), ['TIPLOC', 'LocationName', 'StartDate', 'TimingPointType', 'ZoneResponsible', 'STANOX', 'OffNetwork', 'Quality'], ['Easting', 'Northing'],
                               aggfunc = {'TIPLOC': getCommonPrefix,
                                          'LocationName': getCommonPrefix,
                                          'StartDate': np.min,
                                          'TimingPointType': getCommonPrefix,
                                          'ZoneResponsible': getCommonPrefix,
                                          'STANOX': getCommonPrefix,
                                          'OffNetwork': np.mean,
                                          'Quality': np.min}).reset_index().reset_index()

    LOCsUnique['index'] += 100000
    LOCsUnique.set_index(['index'], inplace = True)
    
    LOCs = LOCs.reset_index().merge(LOCsUnique.reset_index()[['Easting', 'Northing', 'index', 'TIPLOC', 'LocationName']], 'left', ['Easting', 'Northing']).set_index('TIPLOC_x')
    LOCs['index'] = LOCs['index'].astype('Int32')

    PLTs = pd.DataFrame(list(filter(lambda line: line[0] == "PLT", lines)),
                        columns = ['RecordType', 'ActionCode', 'TIPLOC', 'PlatformID', 'StartDate', 'EndDate',
                                   'PlatformLength', 'PowerSupplyType', 'PassengerDOO', 'NonPassengerDOO'])
    
    PLTs.drop(['RecordType', 'ActionCode', 'EndDate', 'PowerSupplyType'], axis = 1, inplace = True)
    PLTs['index_TIPLOC'] = LOCs.loc[PLTs['TIPLOC']]['index'].values
    
    LOCs = LOCs[~LOCs['index'].isna()]
    LOCs.to_csv(os.path.join(path, 'input\\BPLAN_LOC.csv'))
    
    PLTs = PLTs[~PLTs['index_TIPLOC'].isna()]
    PLTs['PlatformID'] = PLTs['PlatformID'].str.upper()
    PLTs['index_PlatformID'] = [sorted(PLTs['PlatformID'].unique()).index(PlatformID) + 1 for PlatformID in PLTs['PlatformID']]
    PLTs['index'] = 1000*PLTs['index_TIPLOC']
    PLTs['TIPLOC_PlatformID'] = PLTs['TIPLOC'] + '_' +  PLTs['PlatformID']
    PLTs.set_index(['TIPLOC_PlatformID'], inplace = True)
    PLTs['StartDate'] = PLTs['StartDate'].astype('datetime64', False)
    PLTs['PlatformLength'] = PLTs['PlatformLength'].replace('', 0)
    PLTs['PlatformLength'] = PLTs['PlatformLength'].astype('int32', False)
    
    for col in ['PassengerDOO', 'NonPassengerDOO']:
        PLTs[col] = [aBool == "Y" for aBool in PLTs[col]]
    
    for col in ['Easting', 'Northing', 'Quality']:
        PLTs[col] = 0

    for i, row in PLTs.iterrows():
        myPlatformNum = re.sub('[^0-9]', '', row['PlatformID'])
        myLOC = LOCs.loc[row['TIPLOC']]
        if myLOC['Quality'] > 0:
            OSMpltData = get_OSM_platform_data(path, row['index_TIPLOC'], f"{myLOC['TIPLOC_y']}: {myLOC['LocationName_y']}", myLOC['Easting'], myLOC['Northing'], 250)
            OSMpltDataFil = OSMpltData[OSMpltData['Platform'] == row['PlatformID']]
            if len(OSMpltDataFil) > 0:
                OSMloc = OSMpltDataFil['Location'].iloc[0]
                PLTs.loc[i, 'Easting'] = OSMloc.x
                PLTs.loc[i, 'Northing'] = OSMloc.y
                PLTs.loc[i, 'Quality'] = 2
                PLTs.loc[i, 'index'] += row['index_PlatformID']
            elif len(myPlatformNum) > 0:
                OSMpltData['Platform'] = [re.sub('[^0-9]', '', str(platform)) for platform in OSMpltData['Platform']]
                OSMpltDataFil = OSMpltData[OSMpltData['Platform'] == myPlatformNum]
                if len(OSMpltDataFil) > 0:
                    OSMloc = OSMpltDataFil['Location'].iloc[0]
                    PLTs.loc[i, 'Easting'] = OSMloc.x
                    PLTs.loc[i, 'Northing'] = OSMloc.y
                    PLTs.loc[i, 'Quality'] = 1
                    PLTs.loc[i, 'index'] += row['index_PlatformID']
    
    PLTs = PLTs[PLTs['Quality'] > 0]
    PLTs.to_csv(os.path.join(path, 'input\\BPLAN_PLT.csv'))

    PLTsUnique = pd.pivot_table(PLTs.reset_index(), ['PlatformID', 'StartDate', 'PlatformLength', 'PassengerDOO', 'NonPassengerDOO', 'index_TIPLOC', 'index_PlatformID', 'Easting', 'Northing', 'Quality'], ['index'],
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

    return LOCsUnique, PLTsUnique

def progressBar(myRange):
    class ProgWin(wx.Frame):
        def __init__(self, parent, title): 
            super(ProgWin, self).__init__(parent, title = title,size = (300, 200))  
            self.InitUI()
        def InitUI(self):    
            self.count = 0 
            pnl = wx.Panel(self)
            self.gauge = wx.Gauge(pnl, range = myRange, size = (300, 25), style =  wx.GA_HORIZONTAL)
            self.SetSize((300, 100)) 
            self.Centre() 
            self.Show(True)
    
    prog = ProgWin(None, 'wx.Gauge')
    return prog

def getVisumLOCs(LOCsUnique, myVer, myShp):
    Visum = com.Dispatch('Visum.Visum.230')
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
    Visum.Net.SetProjection(projString, False)
    Visum.Net.SetAttValue('LeftHandTraffic', 1)
    ImportShapeFilePara = Visum.IO.CreateImportShapeFilePara()
    ImportShapeFilePara.CreateUserDefinedAttributes = True
    ImportShapeFilePara.ObjectType = 0
    ImportShapeFilePara.SetAttValue('Directed', True)
    Visum.IO.ImportShapefile(myShp, ImportShapeFilePara)
    fixDirectedNet(Visum)
    MyMapMatcher = Visum.Net.CreateMapMatcher()
    LinkType = Visum.Net.AddLinkType(2)
    LinkType.SetAttValue('TSysSet', '2')
    for uda, dtype in [['OffNetwork', 2], ['Quality', 1], ['STANOX', 5], ['StartDate', 5], ['TimingPointType', 5], ['ZoneResponsible', 5]]:
        Visum.Net.Stops.AddUserDefinedAttribute(uda, uda, uda, dtype)
    Visum.Graphic.StopDrawing = True
    ex = wx.App()
    prog = progressBar(LOCsUnique.index.max() - 100000)
    for i, row in LOCsUnique.iterrows():
        Node = Visum.Net.AddNode(i, row['Easting'], row['Northing'])
        Node.SetAttValue('Code', row['TIPLOC'])
        Node.SetAttValue('Name', row['LocationName'])
        Stop = Visum.Net.AddStop(i, row['Easting'], row['Northing'])
        Stop.SetAttValue('Code', row['TIPLOC'])
        Stop.SetAttValue('Name', row['LocationName'])
        Stop.SetAttValue('Quality', row['Quality'])
        Stop.SetAttValue('STANOX', row['STANOX'])
        Stop.SetAttValue('StartDate', str(row['StartDate']))
        Stop.SetAttValue('TimingPointType', row['TimingPointType'])
        Stop.SetAttValue('ZoneResponsible', row['ZoneResponsible'])
        StopArea = Visum.Net.AddStopArea(1000*i, i, i, row['Easting'], row['Northing'])
        StopArea.SetAttValue('Code', row['TIPLOC'])
        StopArea.SetAttValue('Name', 'Platform Unknown')
        StopPoint = Visum.Net.AddStopPointOnNode(1000*i, StopArea, i)
        StopPoint.SetAttValue('Code', row['TIPLOC'])
        StopPoint.SetAttValue('Name', 'Platform Unknown')
        unsatis = True
        fil_string = '[TYPENO]=1'
        nTRID = 1
        while unsatis & (nTRID <= 3):
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
                    Visum.Net.Links.ItemByKey(split_Link.Link.AttValue('FromNodeNo'), split_no).SetNo(split_no + 3)
                    Visum.Net.Links.ItemByKey(split_no, split_Link.Link.AttValue('ToNodeNo')).SetNo(split_no + 6)
                    Visum.Net.AddLink(split_no, split_no, i, 2)
                fil_string += f"&[TRID]!=\"{split_TRID}\""
                nTRID += 1
        prog.gauge.SetValue(i - 100000)
    Visum.Graphic.StopDrawing = False
    Visum.Net.Turns.GetFilteredSet('[FromLink\\TypeNo]=2&[ToLink\\TypeNo]=2&[FromLink\\No]!=[ToLink\\No]').SetAllAttValues('TSysSet', '')
    Visum.IO.SaveVersion(myVer)


def getVisumPLTs(PLTsUnique, myPLTsVer, myLOCsVer, LOCsUnique, output):
    Visum = com.Dispatch('Visum.Visum.230')
    Visum.IO.LoadVersion(myLOCsVer)
    Visum.Net.Links.GetFilteredSet('[TypeNo]=1').SetActive()
    Visum.Graphic.StopDrawing = True
    ex = wx.App()
    prog = progressBar(PLTsUnique.index.max() - 100000000)
    for i, row in PLTsUnique.iterrows():
        addStopPoint(Visum, i, row, 250, LOCsUnique)
        prog.gauge.SetValue(i - 100000000)
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

def main():

    path = os.path.dirname(__file__)

    myPickle = os.path.join(path, 'input\\BPLAN.p')
    
    if os.path.exists(myPickle):
        print('Read old already processed BPLAN pickle results from cache')
        with open(myPickle, 'rb') as f:
            LOCsUnique, PLTsUnique = pickle.load(f)
    else:
        print('Reprocessed BPLAN to obtain new pickle results and saved to cache')
        LOCsUnique, PLTsUnique = processBPLAN(path)
        with open(myPickle, 'wb') as f:
            pickle.dump([LOCsUnique, PLTsUnique], f)
    
    myShp = os.path.join(path, 'Shp\\NR_Full_Network.shp')
    myLOCsVer = os.path.join(path, 'output_Visum\\LOCs_Only.ver')
    myPLTsVer = os.path.join(path, 'output_Visum\\LOCs_and_PLTs.ver')
    output = os.path.join(path, 'output_GTFS\\stops.txt')

    if os.path.exists(myLOCsVer):
        print('Read old processed BPLAN LOCs Version file from cache')
    else:
        print('Reprocessed BPLAN to obtain new LOCs Version file and saved to cache')
        getVisumLOCs(LOCsUnique, myLOCsVer, myShp)

    if os.path.exists(myPLTsVer):
        print('Read old processed BPLAN PLTs Version file from cache')
    else:
        print('Reprocessed BPLAN to obtain new PLTs Version file and saved to cache')
        getVisumPLTs(PLTsUnique, myPLTsVer, myLOCsVer, LOCsUnique, output)

    print('Done')

if __name__ == "__main__":
    main()
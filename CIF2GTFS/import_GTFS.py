import os
import pandas as pd
import win32com.client as com

import sys
sys.path.append(os.path.dirname(__file__))

import get_inputs as gi
import CIF2GTFS.results.create_O00 as O00


def createGTFSPuti(gtfsPath, putiFolder):

    template = '''<?xml version="1.0" encoding="UTF-8" standalone="no" ?>
<GTFSIMPORT VERSION="1">

  <IMPORTGTFSPARA AGGREGATELINES="0" CALENDARPERIOD="CALENDARPERIODYEAR" DEFAULTTRANSFERWALKTIME="1min" FILENAME="placeholder" GUESSOPERATINGPERIODS="1" IMPORTSHAPES="1" INSERTMULTIPLESTOPPOINTSFORSTOPSONDIFFERENTSHAPES="0" MAXDISTANCEBETWEENSTOPPOINTANDSHAPE="10m" MAXDWELLTIMEFORPASSTRIPCHAIN="1h"/>

</GTFSIMPORT>
'''
    template = template.replace("placeholder", gtfsPath)

    path = os.path.join(putiFolder, "import_GTFS_to_Visum_23.puti")

    with open(path, 'w') as f:
        f.write(template)


def createPutSupplyPuti(putSupplyPath, putiFolder, tsys_path):

    header = '''<?xml version="1.0" encoding="UTF-8" standalone="no" ?>
<PUTUPDATER VERSION="1">

  <IMPORTPUTSUPPLYBASEPARA IMPORTBLOCKS="0" IMPORTWALKLINKS="0" REPLACEORDELETEACTIVELINES="1" USEONLYACTIVEVEHJOURNEYSECTIONS="0" VERSIONFILENAME="placeholder"/>

  <IMPORTPUTSUPPLYSTOPPOINTPARA HANDLEMISSINGSTOPPOINTCANDIDATE="ASKUSERFORSTOPPOINTSELECTION" INSERTDIRECTEDSTOPPOINTSONLINKS="0" MINDISTANCEBETWEENLINKSTOPPOINTS="100m" SOURCESTOPPOINTNUMBERATTR="SPECIALENTRY_EMPTY" STOPAREAKEYATTRSOURCE="NO" STOPAREAKEYATTRTARGET="NO" STOPKEYATTRSOURCE="NO" STOPKEYATTRTARGET="NO" STOPPOINTKEYATTRSOURCE="NO" STOPPOINTKEYATTRTARGET="NO" STOPPOINTUSAGE="INSERTNODESANDLINKS" USEONLYACTIVENODESLINKSANDSTOPPOINTS="0" USEONLYIDREFERENCEFOREXISTINGSTOPPOINTS="1" USESTOPPOINTKEYATTRIBUTE="1"/>

  <IMPORTPUTSUPPLYROUTINGPARA ADDITIONALTURNRELCOSTFORSHARPANGLES="100" BUNDLELINEROUTES="1" BUNDLELINEROUTESWITHSYSTEMROUTES="1" CANCELSHORTESTPATHSEARCHAFTERCOSTFACTOR="4" CANCELSHORTESTPATHSEARCHAFTERCOSTSUPPLEMENT="100" COSTPROBABILITYTHRESHOLD="1" FREECOORDINATEWEIGHT="0.75" LINEMATCHINGTSYSWEIGHT="2" LINENONMATCHINGTSYSWEIGHT="1.2" LINKCOSTATTR="LENGTH" MATCHINGPOINTSELECTION="StopPoints" MAXNUMBEROFCANDIDATES="1000000000" MAXVALUEOFSHARPANGLE="95" NOSTOPPOINTWEIGHT="0.75" REPLACEDLINEEXACTMATCHWEIGHT="5" REPLACEDLINEMATCHINGTSYSWEIGHT="3" REPLACEDLINENONMATCHINGTSYSWEIGHT="1.2" SAMEFROMANDTOCANDIDATEWEIGHT="0.5" SNAPRADIUS="250m" STANDARDDEVIATION="10m" SYSROUTEWEIGHT="1.1" TURNRELCOSTATTR="0.0" USEADDITIONALTURNRELCOSTFORSHARPANGLES="1" USEDEVIATIONTESTINSHORTESTPATHSEARCH="0" USEONLYOPENLINKS="1" USEONLYOPENTURNRELS="1"/>

  <IMPORTPUTSUPPLYTSYSPARA>'''

    
    footer = '''</IMPORTPUTSUPPLYTSYSPARA>

  <IMPORTPUTSUPPLYNETOBJECTPARA AGGREGATEVALIDDAYS="1" LINKTYPEOFNEWLINKS="99">
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSET" NETOBJECTTYPE="VALIDDAYS" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSET" NETOBJECTTYPE="OPERATINGPERIOD" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSET" NETOBJECTTYPE="VEHUNIT" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSETWITHAGGREGATION" NETOBJECTTYPE="VEHCOMB" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSET" NETOBJECTTYPE="OPERATOR" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSET" NETOBJECTTYPE="STOP" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSET" NETOBJECTTYPE="STOPAREA" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSET" NETOBJECTTYPE="STOPPOINT" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="PREFIX" NETOBJECTTYPE="MAINLINE" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="PREFIX" NETOBJECTTYPE="LINE" OFFSET="0" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSETWITHAGGREGATION" NETOBJECTTYPE="VEHJOURNEY" OFFSET="17211" PREFIX=""/>
    <IMPORTPUTSUPPLYNETOBJECTPARAENTRY CONFLICTAVOIDANCE="OFFSETWITHAGGREGATION" NETOBJECTTYPE="COORDGRP" OFFSET="0" PREFIX=""/>
  </IMPORTPUTSUPPLYNETOBJECTPARA>

</PUTUPDATER>
'''
    header = header.replace("placeholder", putSupplyPath)

    #<IMPORTPUTSUPPLYTSYSPARAENTRY TSYSCODESOURCE="W" TSYSCODETARGET="W"/>
    dfTsys = pd.read_csv(tsys_path, low_memory = False)
    tsys_codes = dfTsys.Code.tolist()

    path = os.path.join(putiFolder, "import_PuT_supply_from_Visum_23.puti")

    with open(path, 'w') as f:
        f.write(header)
        for t in tsys_codes:
            f.write(f'  <IMPORTPUTSUPPLYTSYSPARAENTRY TSYSCODESOURCE="{t}" TSYSCODETARGET="{t}"/>')
        f.write(footer)


def main():

    path = os.path.dirname(__file__)

    proj_string = """
    PROJCS["British_National_Grid_TOWGS",
        GEOGCS["GCS_OSGB_1936",
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
    gtfsPath = os.path.join(path, "output\\GTFS.zip")
    putiPath = os.path.join(path, "puti")
    putSupplyPath = os.path.join(path, 'output\\VISUM\\GTFS_Only.ver')
    tsysPath = os.path.join(path, 'input\\TSys_definitions.csv')
    createGTFSPuti(gtfsPath, putiPath)
    createPutSupplyPuti(putSupplyPath, putiPath, tsysPath)

    print('Importing GTFS to new Visum Version File...')
    Visum = com.Dispatch('Visum.Visum.240')
    Visum.IO.ImportPuTProject(os.path.join(path, 'puti\\import_GTFS_to_Visum_23.puti'))
    Visum.Net.SetProjection(proj_string, False)
    Visum.Net.Stops.SetMultipleAttributes(['No'], Visum.Net.Stops.GetMultipleAttributes(['GTFS_stop_id']))
    Visum.Net.StopAreas.SetMultipleAttributes(['No'], Visum.Net.StopAreas.GetMultipleAttributes(['GTFS_stop_id']))
    Visum.Net.StopPoints.SetMultipleAttributes(['No'], Visum.Net.StopPoints.GetMultipleAttributes(['StopAreaNo']))
    Visum.Net.Directions.SetMultipleAttributes(['Code', 'Name'], (('>', 'Direction: up'), ('<', 'Direction: down')))
    CalendarPeriod_T = Visum.Net.CalendarPeriod.AttValue('Type')
    CalendarPeriod_VF = Visum.Net.CalendarPeriod.AttValue('ValidFrom')
    CalendarPeriod_VU = Visum.Net.CalendarPeriod.AttValue('ValidUntil')
    Visum.IO.SaveVersion(os.path.join(path, 'output\\VISUM\\GTFS_Only.ver'))

    print('Importing GTFS PT supply into prepared Visum network...')
    Visum.IO.LoadVersion(os.path.join(path, 'cached_data\\VISUM\\LOCs_and_PLTs_ZonesConnectorsXferLinks.ver'))
    Visum.Net.CalendarPeriod.SetAttValue('Type', CalendarPeriod_T)
    try:
        Visum.Net.CalendarPeriod.SetAttValue('ValidFrom', CalendarPeriod_VF)
        Visum.Net.CalendarPeriod.SetAttValue('ValidUntil', CalendarPeriod_VU)
        print('Note: Your calendar period is in the past.')
    except:
        Visum.Net.CalendarPeriod.SetAttValue('ValidUntil', CalendarPeriod_VU)
        Visum.Net.CalendarPeriod.SetAttValue('ValidFrom', CalendarPeriod_VF)
        print('Note: Your calendar period is in the future.')
    Visum.Net.Nodes.SetActive()
    Visum.Net.Links.SetActive()
    Visum.Net.StopPoints.SetActive()
    LinkType = Visum.Net.AddLinkType(99)
    LinkType.SetAttValue('TSysSet', '')
    LinkType = None
    Visum.IO.ImportPuTProject(os.path.join(path, 'puti\\import_PuT_supply_from_Visum_23.puti'))
    Visum.Net.TimeProfileItems.AddUserDefinedAttribute('Speed', 'Speed', 'Speed', 15, formula = '3600*[Sum:UsedLineRouteItems\\PostLinkLength]/[PostRunTime]')
    journeyDetails = pd.read_csv(os.path.join(path, 'cached_data\\JourneyDetails.txt'), low_memory = False)
    journeyDetails.drop(['Key', 'JourneyID', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'start_date', 'end_date'], axis = 1, inplace = True)
    VJs = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultiAttValues('Name'), columns = ['No', 'service_id'])
    VJs['service_id'] = VJs['service_id'].str.replace('_trip', '_service')
    journeyDetails = VJs.merge(journeyDetails, 'left', 'service_id')
    journeyDetails.fillna('', inplace = True)
    journeyDetails.drop(columns = ['No', 'service_id'], inplace = True)
    for col in journeyDetails.columns.values:
        Visum.Net.VehicleJourneys.AddUserDefinedAttribute(col, col, col, 5)
    Visum.Net.VehicleJourneys.SetMultipleAttributes(journeyDetails.columns.values, journeyDetails.values)

    O00.Visum = Visum
    O00.main()

    Visum.IO.SaveVersion(os.path.join(path, 'output\\VISUM\\Network+Timetable.ver'))

    print('Done')

if __name__ == "__main__":

    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    importTimetable = gi.readTimetableInputs(input_path)

    if importTimetable[0] == "TRUE":
        main()
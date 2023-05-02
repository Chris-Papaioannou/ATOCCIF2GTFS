import os
import pandas as pd
import win32com.client as com

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

    print('Importing GTFS to new Visum Version File...')
    Visum = com.Dispatch('Visum.Visum.230')
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
    Visum.IO.LoadVersion(os.path.join(path, 'cached_data\\VISUM\\LOCs_and_PLTs.ver'))
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
    LinkType.SetAttValue('TSysSet', '2')
    LinkType = None
    Visum.IO.ImportPuTProject(os.path.join(path, 'puti\\import_PuT_supply_from_Visum_23.puti'))
    Visum.Net.TimeProfileItems.AddUserDefinedAttribute('Speed', 'Speed', 'Speed', 15, formula = '3600*[Sum:UsedLineRouteItems\\PostLinkLength]/[PostRunTime]')
    journeyDetails = pd.read_csv(os.path.join(path, 'cached_data\\JourneyDetails.csv'), low_memory = False)
    journeyDetails.drop(['Key', 'JourneyID', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'start_date', 'end_date'], axis = 1, inplace = True)
    VJs = pd.DataFrame(Visum.Net.VehicleJourneys.GetMultiAttValues('Name'), columns = ['No', 'service_id'])
    VJs['service_id'] = VJs['service_id'].str.replace('_trip', '_service')
    journeyDetails = VJs.merge(journeyDetails, 'left', 'service_id')
    journeyDetails.fillna('', inplace = True)
    journeyDetails.drop(columns = ['No', 'service_id'], inplace = True)
    for col in journeyDetails.columns.values:
        Visum.Net.VehicleJourneys.AddUserDefinedAttribute(col, col, col, 5)
    Visum.Net.VehicleJourneys.SetMultipleAttributes(journeyDetails.columns.values, journeyDetails.values)

    Visum.IO.SaveVersion(os.path.join(path, 'output\\VISUM\\LOCs_and_PLTs_with_GTFS.ver'))

    print('Done')

if __name__ == "__main__":
    main()
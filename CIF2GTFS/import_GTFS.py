import os
import win32com.client as com

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

Visum = com.Dispatch('Visum.Visum.220')
Visum.IO.ImportPuTProject(os.path.join(path, 'import_GTFS.puti'))
Visum.Net.SetProjection(proj_string, False)
Visum.Net.Stops.SetMultipleAttributes(['No'], Visum.Net.Stops.GetMultipleAttributes(['GTFS_stop_id']))
Visum.Net.StopAreas.SetMultipleAttributes(['No'], Visum.Net.StopAreas.GetMultipleAttributes(['GTFS_stop_id']))
Visum.Net.StopPoints.SetMultipleAttributes(['No'], Visum.Net.StopPoints.GetMultipleAttributes(['StopAreaNo']))
Visum.Net.Directions.SetMultipleAttributes(['Code', 'Name'], (('>', 'Direction: up'), ('<', 'Direction: down')))
CP_T = Visum.Net.CalendarPeriod.AttValue('Type')
CP_VF = Visum.Net.CalendarPeriod.AttValue('ValidFrom')
CP_VU = Visum.Net.CalendarPeriod.AttValue('ValidUntil')
Visum.IO.SaveVersion(os.path.join(path, 'output_Visum\\import_GTFS.ver'))
Visum.IO.LoadVersion(os.path.join(path, 'output_Visum\\DetailedNetwork_Processed.ver'))
Visum.Net.CalendarPeriod.SetAttValue('Type', CP_T)
Visum.Net.CalendarPeriod.SetAttValue('ValidFrom', CP_VF)
Visum.Net.CalendarPeriod.SetAttValue('ValidUntil', CP_VU)
Visum.IO.ImportPuTProject(os.path.join(path, 'import_PuT_supply_from_Visum.puti'))
print('done')
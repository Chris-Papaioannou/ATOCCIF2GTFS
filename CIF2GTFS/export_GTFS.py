import os
import pandas as pd
import win32com.client as com
import traceback
import sys
sys.path.append(os.path.dirname(__file__))

import get_inputs as gi

import logging

logging.basicConfig(
    filename="ModelBuilder.log",
    encoding="utf-8",
    filemode="a",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M",
    level=logging.INFO # Change to logging.DEBUG for more details
)


def main():

    try:

        path = os.path.dirname(__file__)
        input_path = os.path.join(path, "input\\inputs.csv")
        
        exportGTFS = gi.readGTFSInputs(input_path)
        exportGTFSbool = exportGTFS[0]
        tsysPath = exportGTFS[1]

        if exportGTFSbool:
            Visum = com.Dispatch("Visum.Visum.240")
            Visum.LoadVersion(os.path.join(path, f"output\\VISUM\\Network+Timetable_MergeStops.ver"))

            runID = gi.getRunID(input_path)

            tsysLookup = pd.read_csv(tsysPath)
            tsysRouteType = dict(zip(tsysLookup.Code, tsysLookup.route_type))

            rtAtt = Visum.Net.TSystems.AddUserDefinedAttribute("route_type", "route_type", "route_type", 1)
            AllTSys = Visum.Net.TSystems.Iterator
            while AllTSys.Valid:
                TSys = AllTSys.Item
                code = TSys.AttValue("CODE")
                TSys.SetAttValue("route_type", tsysRouteType.get(code, -1))
                AllTSys.Next()

            exportParams = Visum.IO.CreateExportGTFSPara()
            exportParams.SetAttValue("AgencyTimeZone", 'Europe/London')
            exportParams.SetAttValue("CalendarFrom", Visum.Net.CalendarPeriod.AttValue('ValidFrom'))
            exportParams.SetAttValue("CalendarTo", Visum.Net.CalendarPeriod.AttValue('ValidUntil'))
            exportParams.SetAttValue("ExportShapeType", 1)
            exportParams.SetAttValue("ExportUDAs", False)
            exportParams.SetAttValue("OperatorUDAForAgencyUrl", 'GTFS_AGENCY_URL')
            exportParams.SetAttValue("Overwrite", True)
            exportParams.SetAttValue("TSysUDAForGTFSMapping", 'route_type')
            exportParams.SetAttValue("ZipFileName", os.path.join(path, 'output', f"{runID}_GTFS.zip"))

            Visum.IO.ExportGTFS(exportParams)
    except:
        logging.error(traceback.format_exc())

if __name__ == '__main__':
    main()
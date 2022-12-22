import sys
import win32com.client as com

#argv1 = sys.argv[1]

Visum = com.Dispatch("Visum.Visum.220")
Visum.IO.ImportPuTProject('import_GTFS.puti')
Visum.IO.SaveVersion('test.ver')
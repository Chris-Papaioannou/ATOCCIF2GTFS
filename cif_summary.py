import pandas as pd

cif_path = r"C:\Users\david.aspital\PTV Group\Team Network Model T2BAU - General\07 Model Files\07 May23 Timetable Reimport\ATOCCIF2GTFS\CIF2GTFS\input\May23 Full CIF 230413.CIF"

with open(cif_path, "r") as file:
    cif = file.read()



# Count new, revised and deleted services
bsn = cif.count("\nBSN")
bsr = cif.count("\nBSR")
bsd = cif.count("\nBSD")


with open(r"CIF_Summary.txt", "w+") as file:
    file.write(f"CIF File - {cif_path}\n")
    file.write(f"New services - {bsn}\n")
    file.write(f"Revised services - {bsr}\n")
    file.write(f"Deleted services - {bsd}\n")
    file.write("\n")
    #tiplocs_lines.to_csv(file, index=False, sep="\t", line_terminator="\n")

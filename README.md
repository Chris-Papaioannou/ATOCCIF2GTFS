# ATOCCIF2GTFS2VISUM

The puropose of this tool is to convert a ATOC CIF format timetable (commonly used in Great Britain) to a platform-resolution GTFS format, and then import it into an automatically generated [PTV Visum](https://www.myptv.com/en/mobility-software/ptv-visum) network version file.

## Disclaimer

ATOC CIF is a complicated format and the conversions in this tool are not perfect.

## Usage

The tool is written in C# .NET Core 3.1 & Python 3.9.9. It was developed and tested in Visual Studio Code February 2023 (version 1.76.1) for Windows. There are ways to open the solution and compile the project on Mac and Linux, either using Visual Studio or the [.NET Core Runtime or SDK](https://dotnet.microsoft.com/download) but I have not tested them.

## Inputs

### Network

Shapefile provided by [Network Rail](https://www.networkrail.co.uk/). Expected fields include:
['OBJECTID', 'ASSETID', 'L_LINK_ID', 'L_SYSTEM', 'L_VAL', 'L_QUALITY', 'ELR', 'TRID',
 'TRCODE', 'L_M_FROM', 'L_M_TO', 'VERSION_NU', 'VERSION_DA', 'SOURCE', 'EDIT_STATU',
 'IDENTIFIED', 'TRACK_STAT', 'LAST_EDITE', 'LAST_EDI_1', 'CHECKED_BY', 'CHECKED_DA',
 'VALIDATED_', 'VALIDATED1', 'EDIT_NOTES', 'PROIRITY_A', 'SHAPE_LENG', 'TRID_CAT']

However, only [TRCODE]() and [TRID]() are used in the current implementation.

An older branch of the tool used to also use [ELR] to apply more refined user-defined link filtering when snapping platform locations to the network.

However, the tool has now transitioned to an [OpenStreetMap](https://www.openstreetmap.org/about) based approach using the [Overpass API](https://python-overpy.readthedocs.io/en/latest/).
 
### BPLAN

This tool uses [Location (LOC) records] & [Platforms and Sidings (PLT) records] from [BPLAN](https://wiki.openraildata.com/index.php?title=BPLAN_data_structure) to determine what possible TIPLOC codes and PlatformIDs could be expected in any given CIF file. This reduces the regularity of updates required to the underlying network. BPLAN is updated every 6 months and can be downloaded [here](https://wiki.openraildata.com/index.php?title=BPLAN_Geography_Data).

### NaPTAN

As the Eastings and Northings contained withing [BPLAN](https://wiki.openraildata.com/index.php?title=BPLAN_data_structure) are notoriously unreliable, this program infills Eastings and Northings from the [RLY] subset of [NaPTAN] when available (i.e. nearly all passenger rail station TIPLOCs) using the national Stops.csv file. This can be downloaded [here](https://beta-naptan.dft.gov.uk/download). An up to date version should be used or newer stations risk being ommitted from the output file.

It is anticipated that in the future, this program could be updated to use platform locations from the [RPL] subset of NaPTANs also. However, currently only light rail platforms from the [PLT] subset have significant coverage. Therefore the tool currently only uses the [OpenStreetMap](https://www.openstreetmap.org/about) based approach using the [Overpass API](https://python-overpy.readthedocs.io/en/latest/) described.

### ATOC CIF

The ATOC CIF tested is a national passenger rail file provided by [Network Rail](https://www.networkrail.co.uk/). If the user wishes to update the timetable, then railway timetables for Great Britain are available from [The Rail Delivery Group](http://data.atoc.org/). You will need to create an account to download the data, which is available for free, and licensed under [The Creative Commons Attribution 2.0 England and Wales license]( https://creativecommons.org/licenses/by/2.0/uk/legalcode). This permits sharing of the original timetable file and its derivative version (the GTFS version of that timetable) while recognising its origin, as above.

## Process

## Warnings

### Prio. = High

### Prio. = Low

## Cached Data

### OSM Platform Images

### OSM Platform Pickles

### BPLAN Pickle

### BPLAN CSVs

### Network Visum Version Files

### Network Build Log Files

## Outputs

### Zipped GTFS Files

### Output Visum Version Files

### GTFS Import Log Files

## Guide to editing OpenStreetMap

## License

This code is released under the MIT License, as included in this repository.
Both timetables, the original in ATOC CIF format and the derviative work in GTFS format are provided under [The Creative Commons Attribution 2.0 England and Wales license]( https://creativecommons.org/licenses/by/2.0/uk/legalcode) as required by the original data provider. This license includes the specific clause that "You must not sublicense the Work". All usage should acknowledge both this repository and the original data source.

## Thanks

This project is supported indirectly (and with no guarantee or liability) by partners of [PTV Group](https://company.ptvgroup.com/en/) including [Network Rail](https://www.networkrail.co.uk/). The parsing of the cif file and some of the GTFS file formatting draw heavily from [ATOCCIF2GTFS](https://github.com/odileeds/ATOCCIF2GTFS) shared on GitHub by [ODI Leeds](https://github.com/odileeds) and written by [Thomas Forth](https://github.com/thomasforth), so special thanks to them also.

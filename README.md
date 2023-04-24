# ATOCCIF2GTFS2VISUM

The puropose of this tool is to convert an [ATOC CIF](https://wiki.openraildata.com/index.php?title=CIF_Schedule_Records) format timetable (commonly used in Great Britain) to a platform-resolution [GTFS](https://gtfs.org/schedule/reference/) format, and then import it into an automatically generated [PTV Visum](https://www.myptv.com/en/mobility-software/ptv-visum) network version file.

## Disclaimer

[ATOC CIF](https://wiki.openraildata.com/index.php?title=CIF_Schedule_Records) is a complicated format and the conversions in this tool are not perfect.

## Usage

The tool is written in **C#** & **Python**. It was developed and tested in **Visual Studio Code February 2023 (version 1.76.1)**. There are ways to open the solution and compile the project on Mac and Linux, using Visual Studio, .NET Core Runtime, or SDK. However, they have not been tested.

### C# Setup

This has been tested using [.NET SDK 7.0.101 for Windows x64](https://dotnet.microsoft.com/en-us/download/dotnet/7.0). However, any .NET implementation 3.0 through 7.X should work. If you are using Visual Studio Code, you should download and install the [C# for Visual Studio Code (powered by OmniSharp)](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.csharp) extension to enable syntax advice, colour formatting, and debugging for C#.

### Python Setup

This has been tested using the **Python 3.9.9** installation included with **PTV Visum 2023**. However, any implementation of Python 3.9 should work. The user should ensure that their default python installation location for python terminals when debugging .py files directly is the same as the one called by the C# function **ExecProcess** to make sure behaviour is the same. The user will need to **pip install** the following libraries if not already available:

bng_latlon, geopandas, json, matplotlib, numpy, os, overpy, pandas, pickle, re, scipy, shapely, time, win32com, wx

## Inputs

### Network

The shapefile is provided by [Network Rail](https://www.networkrail.co.uk/). Expected fields include:

['OBJECTID', 'ASSETID', 'L_LINK_ID', 'L_SYSTEM', 'L_VAL', 'L_QUALITY', 'ELR', 'TRID',
 'TRCODE', 'L_M_FROM', 'L_M_TO', 'VERSION_NU', 'VERSION_DA', 'SOURCE', 'EDIT_STATU',
 'IDENTIFIED', 'TRACK_STAT', 'LAST_EDITE', 'LAST_EDI_1', 'CHECKED_BY', 'CHECKED_DA',
 'VALIDATED_', 'VALIDATED1', 'EDIT_NOTES', 'PROIRITY_A', 'SHAPE_LENG', 'TRID_CAT']

Only **TRID**, **TRCODE**, and **TRACK_STAT** are used in the current implementation. An older branch of the tool used to also use **ELR** to apply more refined user-defined link filtering when snapping platform locations to the network. However, the tool has now transitioned to an [OpenStreetMap](https://www.openstreetmap.org/about) based approach using the [Overpass API](https://python-overpy.readthedocs.io/en/latest/) (with better examples in the documentation [here](https://wiki.openstreetmap.org/wiki/Overpass_API/Overpass_QL)).
 
### BPLAN

This tool uses **Location (LOC) records** & **Platforms and Sidings (PLT) records** from [BPLAN](https://wiki.openraildata.com/index.php?title=BPLAN_data_structure) to determine what possible **TIPLOC** and **PlatformID** codes could be expected in any given [ATOC CIF](https://wiki.openraildata.com/index.php?title=CIF_Schedule_Records) file. This reduces the regularity of updates required to the underlying network. [BPLAN](https://wiki.openraildata.com/index.php?title=BPLAN_data_structure) is updated every 6 months and can be downloaded [here](https://wiki.openraildata.com/index.php?title=BPLAN_Geography_Data).

### Tiploc Public Export JSON

As the Eastings and Northings contained within [BPLAN](https://wiki.openraildata.com/index.php?title=BPLAN_data_structure) are notoriously unreliable, this program instead converts latitudes and longitudes from a 3rd party spatial source to Eastings and Northings. This JSON can be downloaded from [RailMap](https://railmap.azurewebsites.net/Downloads/). 

It is anticipated that in the future, this program could be updated to use platform locations from the **RPL** subset of the national **Stops.csv** file of **NaPTAN** also. This can be downloaded [here](https://beta-naptan.dft.gov.uk/download). An up-to-date version should be used or newer stations risk being ommitted from the output file. However, currently only light rail platforms from the **PLT** subset have significant coverage. Therefore the tool currently only uses the [OpenStreetMap](https://www.openstreetmap.org/about) based approach using the [Overpass API](https://python-overpy.readthedocs.io/en/latest/) described.

### ATOC CIF

The ATOC CIF tested is a national passenger rail file provided by [Network Rail](https://www.networkrail.co.uk/). If the user wishes to update the timetable, then railway timetables for Great Britain are available from [The Rail Delivery Group](http://data.atoc.org/). You will need to create an account to download the data, which is available for free, and licensed under [The Creative Commons Attribution 2.0 England and Wales license]( https://creativecommons.org/licenses/by/2.0/uk/legalcode). This permits sharing of the original timetable file and its derivative version (the [GTFS](https://gtfs.org/schedule/reference/) version of that timetable) while recognising its origin, as above.

## Process

### 1. Run prepare_network.py

#### a. Process Tiploc Public Export JSON Stop information

The json file is read and converted to a Pandas DataFrame, filtering any TIPLOCs with a character length > 7 (this removes non-Tiploc locations that could never be referenced by any ATOC CIF file). CRS is extracted from a detailed dictionary if available and latitude / longitude is converted to BGS. The dataframe is pivoted on easting and northing to merge any TIPLOCs with the exact same location into one Visum stop. Tiplocs, Stanox, and CRS are concatenated with values separated by commas, whereas the Name field is reduced to the longest possible common prefix umongst all non-blank values.

#### b. Process TBPLAN Platform information

PLT lines are read from BPLAN to get a list of all possible timetabled platform and siding locations and drop unnecessary columns.All platform references are converted to upper case and "Y" / "N" fields are re-cast as booleans. The resultant Pandas DataFrame of platform locations is iterated through, undertaing an OverpassQL query of OSM data for any station not already present in cached data. This searches for all platforms within a 500m by 500m bounding box centred on the Stop location.

#### c. Get Visum LOCs

A new instance of Visum is opened and projection set to BNG with left hand traffic. The input shapefile is read in as a directed network, and the directions are corrected. A Pandas DataFrame of all unique locations from the processed Tiploc Public Export JSON is iterated through, creating one Node, Stop, Stop Area, and Stop Point for each. Each Node is linked to the network with up to 10 dummy links (the closest open unique TRID values up to a distance of 250m). Only U-Turns are allowed at the Stop node itself, so services must rejoin the same TRID they originally left.

#### d. Get Visum PLTs

### 2. Run main C# Process

#### a. Load stops & platforms from cached data to dictionaries

#### b. Read and parse the ATOC CIF timetable file

#### c. Create initial GTFS output

### 3. Run drop_single_stop_trips.py

### 4. Create final zipped GTFS output in C#

### 5. Run import_GTFS.py

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

Users are encouraged to add new platform locations by making contributions to OSM, which can be done through the browser based Edit GUI. The preferred method is to add a new area or edit an existing area. When a new area is drawn, the feature type "Train Platform" should be selected. This can usually be seen automatically if you are creating a new area in proximity of other rail objects. However, in some cases you may need to use the search function.

This will populate nearly all the tag fields you need to fill automatically. Add the platform code from BPLAN to the "Platform Number" field (this should automatically populate the tag "ref" when you save). Just the BPLAN code on its own is preferred. e.g. "4C". However, "Platform 4C" and "Platforms 4C" will also be trimmed appropriately.

If you want multiple platform codes to map to the same OSM platform, they should be separated with semicolons ";". However, ":" and "&" are also used in some locations. It is preferential to split platforms sharing one platform island into 2 separate OSM areas, to ensure stop points in different track directions snap to the correct locations.

## License

This code is released under the **MIT License**, as included in this repository.
Both timetables, the original in [ATOC CIF](https://wiki.openraildata.com/index.php?title=CIF_Schedule_Records) format and the derviative work in [GTFS](https://gtfs.org/schedule/reference/) format are provided under [The Creative Commons Attribution 2.0 England and Wales license]( https://creativecommons.org/licenses/by/2.0/uk/legalcode) as required by the original data provider. This license includes the specific clause that "You must not sublicense the Work". All usage should acknowledge both this repository and the original data source.

## Thanks

This project is supported indirectly (and with no guarantee or liability) by partners of [PTV Group](https://company.ptvgroup.com/en/) including [Network Rail](https://www.networkrail.co.uk/). The parsing of the [ATOC CIF](https://wiki.openraildata.com/index.php?title=CIF_Schedule_Records) file and some of the [GTFS](https://gtfs.org/schedule/reference/) file formatting draw heavily from [ATOCCIF2GTFS](https://github.com/odileeds/ATOCCIF2GTFS) shared on GitHub by [ODI Leeds](https://github.com/odileeds) and written by [Thomas Forth](https://github.com/thomasforth). The accurate positioning of TIPLOC locations would have been much more challenging without the work of **Liam Crozier** of , so thanks to them also.

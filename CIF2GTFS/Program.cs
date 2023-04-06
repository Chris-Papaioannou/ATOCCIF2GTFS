using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace CIF2GTFS
{
    class Program
    {
        static void Main(string[] args)
        {   
            Console.WriteLine("Preparing Visum network...");
            ExecProcess("prepare_network.py");
            
            Console.WriteLine("Loading BPLAN PLTs...");
            List<BPLAN_PLT> PLTs = new List<BPLAN_PLT>();
            using (TextReader textReader = File.OpenText(@"cached_data/BPLAN/PLTs.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader, CultureInfo.InvariantCulture);
                csvReader.Configuration.Delimiter = ",";
                PLTs = csvReader.GetRecords<BPLAN_PLT>().ToList();
            }

            Console.WriteLine("Loading GTFS_STOP_ID keyed dictionary of BPLAN PLTs...");
            Dictionary<string, BPLAN_PLT> PLTsDictionary = new Dictionary<string, BPLAN_PLT>();
            List<GTFSattStop> GTFS_PLTsList = new List<GTFSattStop>();
            foreach (BPLAN_PLT BPLAN_PLT in PLTs)
            {
                if (!PLTsDictionary.ContainsKey(BPLAN_PLT.TIPLOC_PlatformID))
                {
                    PLTsDictionary.Add(BPLAN_PLT.TIPLOC_PlatformID, BPLAN_PLT);

                    GTFSattStop gTFSattStop = new GTFSattStop()
                    {
                        stop_id = BPLAN_PLT.index,
                        stop_code = BPLAN_PLT.PlatformID,
                        stop_name = "Platform " + BPLAN_PLT.PlatformID,
                        location_type = 1
                    };
                    GTFS_PLTsList.Add(gTFSattStop);
                }
            }

            Console.WriteLine("Loading BPLAN LOCs...");
            List<BPLAN_LOC> LOCs = new List<BPLAN_LOC>();
            using (TextReader textReader = File.OpenText("cached_data/BPLAN/LOCs.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader, CultureInfo.InvariantCulture);
                csvReader.Configuration.Delimiter = ",";
                LOCs = csvReader.GetRecords<BPLAN_LOC>().ToList();
            }

            Console.WriteLine("Loading GTFS_STOP_ID keyed dictionary of BPLAN LOCs...");
            Dictionary<string, BPLAN_LOC> LOCsDictionary = new Dictionary<string, BPLAN_LOC>();
            List<GTFSattStop> GTFS_LOCsList = new List<GTFSattStop>();
            foreach (BPLAN_LOC BPLAN_LOC in LOCs)
            {
                if (!LOCsDictionary.ContainsKey(BPLAN_LOC.TIPLOC_x))
                {
                    LOCsDictionary.Add(BPLAN_LOC.TIPLOC_x, BPLAN_LOC);

                    GTFSattStop gTFSattStop = new GTFSattStop()
                    {
                        stop_id = BPLAN_LOC.index,
                        stop_code = BPLAN_LOC.TIPLOC_y,
                        stop_name = BPLAN_LOC.LocationName_y,
                        location_type = 1
                    };
                    GTFS_LOCsList.Add(gTFSattStop);
                }
            }

            Console.WriteLine("Reading the timetable file...");
            List<string> TimetableFileLines = new List<string>(File.ReadAllLines("input/KentMay23.CIF"));
            Dictionary<string, List<StationStop>> StopTimesForJourneyIDDictionary = new Dictionary<string, List<StationStop>>();
            Dictionary<string, JourneyDetail> JourneyDetailsForJourneyIDDictionary = new Dictionary<string, JourneyDetail>();
            string CurrentJourneyID = "";
            string CurrentOperatorCode = "";
            string CurrentTrainType = "";
            string CurrentTrainClass = "";
            string CurrentTrainMaxSpeed = "";
            Calendar CurrentCalendar = null;
            List<string> PrioList = new List<string>();
            foreach (string TimetableLine in TimetableFileLines)
            {
                if (TimetableLine.StartsWith("BS"))
                {
                    CurrentJourneyID = TimetableLine.Substring(2, 7);
                    string StartDateString = TimetableLine.Substring(9, 6);
                    string EndDateString = TimetableLine.Substring(15, 6);
                    string DaysOfOperationString = TimetableLine.Substring(21, 7);
                    // Since a single timetable can have a single Journey ID that is valid at different non-overlapping times a unique Journey ID includes the Date strings and the character at position 79.
                    CurrentJourneyID = CurrentJourneyID + StartDateString + EndDateString;
                    CurrentCalendar = new Calendar()
                    {
                        start_date = "20" + StartDateString,
                        end_date = 2000 + Math.Min(int.Parse(EndDateString.Substring(0, 2)), int.Parse(StartDateString.Substring(0, 2)) + 49) + EndDateString.Substring(2, 4),
                        service_id = CurrentJourneyID + "_service",
                        monday = int.Parse(DaysOfOperationString.Substring(0, 1)),
                        tuesday = int.Parse(DaysOfOperationString.Substring(1, 1)),
                        wednesday = int.Parse(DaysOfOperationString.Substring(2, 1)),
                        thursday = int.Parse(DaysOfOperationString.Substring(3, 1)),
                        friday = int.Parse(DaysOfOperationString.Substring(4, 1)),
                        saturday = int.Parse(DaysOfOperationString.Substring(5, 1)),
                        sunday = int.Parse(DaysOfOperationString.Substring(6, 1))
                    };
                    CurrentTrainType = TimetableLine.Substring(50, 3);
                    CurrentTrainClass = TimetableLine.Substring(53, 3);
                    CurrentTrainMaxSpeed = TimetableLine.Substring(57, 3);
                }
                if (TimetableLine.StartsWith("BX"))
                {
                    CurrentOperatorCode = TimetableLine.Substring(11, 2);
                    JourneyDetail journeyDetail = new JourneyDetail()
                    {
                        JourneyID = CurrentJourneyID,
                        OperatorCode = CurrentOperatorCode,
                        OperationsCalendar = CurrentCalendar,
                        TrainClass = CurrentTrainClass,
                        TrainMaxSpeed = CurrentTrainMaxSpeed,
                        TrainType = CurrentTrainType
                    };
                    JourneyDetailsForJourneyIDDictionary.Add(CurrentJourneyID, journeyDetail);
                }
                if (TimetableLine.StartsWith("LO") || TimetableLine.StartsWith("LI") || TimetableLine.StartsWith("LT"))
                {
                    string thirdSlot = TimetableLine.Substring(10, 4).Trim();
                    string fourthSlot = TimetableLine.Substring(15, 4).Trim();
                    StationStop stationStop = new StationStop()
                    {
                        RecordIdentity = TimetableLine.Substring(0, 2).Trim(),
                        Location = TimetableLine.Substring(2, 7).Trim()
                    };
                    if (TimetableLine.StartsWith("LI"))
                    {
                        string SAT = TimetableLine.Substring(10, 5).Trim();
                        string SDT = TimetableLine.Substring(15, 5).Trim();
                        string SP = TimetableLine.Substring(20, 5).Trim();
                        if (SAT != "" && SDT != "")
                        {
                            stationStop.ScheduledArrivalTime = stringToTimeSpan(SAT);
                            stationStop.ScheduledDepartureTime = stringToTimeSpan(SDT);
                            stationStop.pudoType = 0;
                        }
                        else
                        {
                            stationStop.ScheduledArrivalTime = stringToTimeSpan(SP);
                            stationStop.ScheduledDepartureTime = stringToTimeSpan(SP);
                            stationStop.pudoType = 1;
                        }
                        stationStop.Platform = TimetableLine.Substring(33, 3).Trim();
                        stationStop.Line = TimetableLine.Substring(36, 3).Trim();
                    }
                    else
                    {
                        string SDT = TimetableLine.Substring(10, 5).Trim();
                        stationStop.ScheduledArrivalTime = stringToTimeSpan(SDT);
                        stationStop.ScheduledDepartureTime = stringToTimeSpan(SDT);
                        stationStop.pudoType = 0;
                        stationStop.Platform = TimetableLine.Substring(19, 3).Trim();
                        stationStop.Line = TimetableLine.Substring(22, 3).Trim();
                    }
                    if (PLTsDictionary.ContainsKey(stationStop.Location + "_" + stationStop.Platform))
                    {
                        stationStop.PLT = PLTsDictionary[stationStop.Location + "_" + stationStop.Platform];
                        if (StopTimesForJourneyIDDictionary.ContainsKey(CurrentJourneyID))
                        {
                            List<StationStop> UpdatedStationStops = StopTimesForJourneyIDDictionary[CurrentJourneyID];
                            UpdatedStationStops.Add(stationStop);
                            StopTimesForJourneyIDDictionary.Remove(CurrentJourneyID);
                            StopTimesForJourneyIDDictionary.Add(CurrentJourneyID, UpdatedStationStops);
                        }
                        else
                        {
                            StopTimesForJourneyIDDictionary.Add(CurrentJourneyID, new List<StationStop>() { stationStop });
                        }
                    }
                    else if (LOCsDictionary.ContainsKey(stationStop.Location))
                    {
                        stationStop.LOC = LOCsDictionary[stationStop.Location];
                        if (stationStop.Platform != "")
                        {
                            string myWarning = "WARNING (Prio. = Low): " + stationStop.Location + " Platform " + stationStop.Platform + " not found in OSM. Assigned as Platform Unknown instead.";
                            if (!PrioList.Contains(myWarning))
                            {
                                PrioList.Add(myWarning);
                                Console.WriteLine(myWarning);
                            }
                        }
                        if (StopTimesForJourneyIDDictionary.ContainsKey(CurrentJourneyID))
                        {
                            List<StationStop> UpdatedStationStops = StopTimesForJourneyIDDictionary[CurrentJourneyID];
                            UpdatedStationStops.Add(stationStop);
                            StopTimesForJourneyIDDictionary.Remove(CurrentJourneyID);
                            StopTimesForJourneyIDDictionary.Add(CurrentJourneyID, UpdatedStationStops);
                        }
                        else
                        {
                            StopTimesForJourneyIDDictionary.Add(CurrentJourneyID, new List<StationStop>() { stationStop });
                        }
                    }
                    else
                    {
                        string myWarning = "WARNING (Prio. = High): " + stationStop.Location + " skipped as not found in filtered BPLAN.";
                        if (!PrioList.Contains(myWarning))
                        {
                            PrioList.Add(myWarning);
                            Console.WriteLine(myWarning);
                        }
                    }
                }
            }
            Console.WriteLine($"Read {StopTimesForJourneyIDDictionary.Keys.Count} journeys.");
            Console.WriteLine("Creating GTFS output.");
            List<string> Agencies = JourneyDetailsForJourneyIDDictionary.Values.Select(x => x.OperatorCode).Distinct().ToList();
            // AgencyList will hold the GTFS agency.txt file contents
            List<Agency> AgencyList = new List<Agency>();
            // Get all unique agencies from our output
            foreach (string agency in Agencies)
            {
                Agency NewAgency = new Agency()
                {
                    agency_id = agency,
                    agency_name = agency,
                    agency_url = "https://www.google.com/search?q=" + agency + "%20rail%20operator%20code", // google plus name of agency by default
                    agency_timezone = "Europe/London" // Europe/London by default
                };
                AgencyList.Add(NewAgency);
            }
            List<Route> RoutesList = new List<Route>();
            foreach (string journeyID in JourneyDetailsForJourneyIDDictionary.Keys)
            {
                JourneyDetail journeyDetail = JourneyDetailsForJourneyIDDictionary[journeyID];
                Route route = new Route()
                {
                    agency_id = journeyDetail.OperatorCode,
                    route_id = journeyDetail.JourneyID + "_route",
                    route_type = "2",
                    route_short_name = journeyDetail.OperatorCode + "_" + journeyDetail.JourneyID
                };
                RoutesList.Add(route);
            }
            List<Trip> tripList = new List<Trip>();
            foreach (JourneyDetail journeyDetail in JourneyDetailsForJourneyIDDictionary.Values)
            {
                Trip trip = new Trip()
                {
                    route_id = journeyDetail.JourneyID + "_route",
                    service_id = journeyDetail.JourneyID + "_service",
                    trip_id = journeyDetail.JourneyID + "_trip"
                };
                tripList.Add(trip);
            }
            // This export line is more complicated than it might at first seem sensible to be because of an understandable quirk in the GTFS format.
            // Stop times are only given as a time of day, and not a datetime. This causes problems when a service runs over midnight.
            // To fix this we express stop times on a service that started the previous day with times such as 24:12 instead of 00:12 and 25:20 instead of 01:20.
            // I assume that no journey runs into a third day.
            List<StopTime> stopTimesList = new List<StopTime>();
            foreach (string JourneyID in StopTimesForJourneyIDDictionary.Keys)
            {
                List<StationStop> StationStops = StopTimesForJourneyIDDictionary[JourneyID];
                int count = 1;
                bool JourneyStartedYesterdayFlagA = false;
                bool JourneyStartedYesterdayFlagD = false;
                TimeSpan PreviousStopDepartureTime = new TimeSpan(0);
                foreach (StationStop stationStop in StationStops)
                {
                    if (stationStop.ScheduledArrivalTime < PreviousStopDepartureTime)
                    {
                        JourneyStartedYesterdayFlagA = true;
                    }
                    if (stationStop.ScheduledDepartureTime < PreviousStopDepartureTime)
                    {
                        JourneyStartedYesterdayFlagD = true;
                    }
                    int myStop = 0;
                    if (stationStop.PLT != null)
                    {
                        myStop = stationStop.PLT.index;
                    }
                    else
                    {
                        myStop = 1000*stationStop.LOC.index;
                    }
                    StopTime stopTime = new StopTime()
                    {
                        trip_id = JourneyID + "_trip",
                        stop_id = myStop,
                        stop_sequence = count,
                        pickup_type = stationStop.pudoType,
                        drop_off_type = stationStop.pudoType
                    };
                    if (JourneyStartedYesterdayFlagA == true)
                    {
                        stationStop.ScheduledArrivalTime = stationStop.ScheduledArrivalTime.Add(new TimeSpan(24, 0, 0));
                        stopTime.arrival_time = Math.Floor(stationStop.ScheduledArrivalTime.TotalHours).ToString() + stationStop.ScheduledArrivalTime.ToString(@"hh\:mm\:ss").Substring(2, 6);
                    }
                    else
                    {
                        stopTime.arrival_time = stationStop.ScheduledArrivalTime.ToString(@"hh\:mm\:ss");
                    }
                    if (JourneyStartedYesterdayFlagD == true)
                    {
                        stationStop.ScheduledDepartureTime = stationStop.ScheduledDepartureTime.Add(new TimeSpan(24, 0, 0));
                        stopTime.departure_time = Math.Floor(stationStop.ScheduledDepartureTime.TotalHours).ToString() + stationStop.ScheduledDepartureTime.ToString(@"hh\:mm\:ss").Substring(2, 6);
                    }
                    else
                    {
                        stopTime.departure_time = stationStop.ScheduledDepartureTime.ToString(@"hh\:mm\:ss");
                    }
                    stopTimesList.Add(stopTime);
                    PreviousStopDepartureTime = stationStop.ScheduledDepartureTime;
                    count++;
                }
            }
            List<Calendar> calendarList = JourneyDetailsForJourneyIDDictionary.Values.Select(x => x.OperationsCalendar).ToList();
            // write GTFS txts.
            // agency.txt, calendar.txt, calendar_dates.txt, routes.txt, stop_times.txt, stops.txt, trips.txt

            Console.WriteLine("Writing agency.txt");
            TextWriter agencyTextWriter = File.CreateText(@"output/GTFS/agency.txt");
            CsvWriter agencyCSVwriter = new CsvWriter(agencyTextWriter, CultureInfo.InvariantCulture);
            agencyCSVwriter.WriteRecords(AgencyList);
            agencyTextWriter.Dispose();
            agencyCSVwriter.Dispose();

            Console.WriteLine("Writing routes.txt");
            TextWriter routesTextWriter = File.CreateText(@"output/GTFS/routes.txt");
            CsvWriter routesCSVwriter = new CsvWriter(routesTextWriter, CultureInfo.InvariantCulture);
            routesCSVwriter.WriteRecords(RoutesList);
            routesTextWriter.Dispose();
            routesCSVwriter.Dispose();

            Console.WriteLine("Writing trips.txt");
            TextWriter tripsTextWriter = File.CreateText(@"output/GTFS/trips.txt");
            CsvWriter tripsCSVwriter = new CsvWriter(tripsTextWriter, CultureInfo.InvariantCulture);
            tripsCSVwriter.WriteRecords(tripList);
            tripsTextWriter.Dispose();
            tripsCSVwriter.Dispose();

            Console.WriteLine("Writing calendar.txt");
            TextWriter calendarTextWriter = File.CreateText(@"output/GTFS/calendar.txt");
            CsvWriter calendarCSVwriter = new CsvWriter(calendarTextWriter, CultureInfo.InvariantCulture);
            calendarCSVwriter.WriteRecords(calendarList);
            calendarTextWriter.Dispose();
            calendarCSVwriter.Dispose();

            Console.WriteLine("Writing stop_times.txt");
            TextWriter stopTimeTextWriter = File.CreateText(@"cached_data/STOP_TIMES/full.txt");
            CsvWriter stopTimeCSVwriter = new CsvWriter(stopTimeTextWriter, CultureInfo.InvariantCulture);
            stopTimeCSVwriter.WriteRecords(stopTimesList);
            stopTimeTextWriter.Dispose();
            stopTimeCSVwriter.Dispose();

            Console.WriteLine("Dropping trip IDs with only one matched stop from stop_times.txt");
            ExecProcess("drop_single_stop_trips.py");

            Console.WriteLine("Creating a GTFS .zip file");
            if (File.Exists(@"output/GTFS.zip"))
            {
                File.Delete(@"output/GTFS.zip");
            }
            ZipFile.CreateFromDirectory(@"output/GTFS", @"output/GTFS.zip", CompressionLevel.Optimal, false, Encoding.UTF8);

            Console.WriteLine("Importing GTFS to Visum...");
            ExecProcess("import_GTFS.py");

            Console.WriteLine("Done");
        }

        static void ExecProcess(string my_script)

        // This allows you to use C# as a Shell to run a Python Process
        {

            // 1) Create Process Info
            var psi = new ProcessStartInfo();
            psi.FileName = @"C:\Program Files\PTV Vision\PTV Visum 2023\Exe\Python\python.exe";

            // 2) Provide script and arguments
            var script = my_script;
            var ver_path = @"C:\Users\PLACEHOLDER";
            psi.Arguments = $"\"{script}\" \"{ver_path}\"";

            // 3) Process configuration
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;

            // 4) Execute process and get output
            var errors = "";
            var results = "";

            using(var process = Process.Start(psi))
            {
                errors = process.StandardError.ReadToEnd();
                results = process.StandardOutput.ReadToEnd();
            }

            // 5) Display output
            Console.WriteLine("ERRORS:");
            Console.WriteLine(errors);
            Console.WriteLine("Results:");
            Console.WriteLine(results);

        }

        static TimeSpan stringToTimeSpan(string input)
        {
            // input is expected to be HHMM
            int hours = int.Parse(input.Substring(0, 2));
            int minutes = int.Parse(input.Substring(2, 2));
            if (input.EndsWith("H"))
            {
                TimeSpan timeSpan = new TimeSpan(hours, minutes, 30);
                return timeSpan;
            }
            else
            {
                TimeSpan timeSpan = new TimeSpan(hours, minutes, 0);
                return timeSpan;
            }
        }
    }

    public class JourneyDetail
    {
        public string JourneyID { get; set; }
        public string OperatorCode { get; set; }
        public string TrainType { get; set;}
        public string TrainClass { get; set;}
        public string TrainMaxSpeed { get; set;}
        public Calendar OperationsCalendar { get; set; }
    }

    public class StationStop
    {
        public string RecordIdentity { get; set; }
        public string Location { get; set; }
        public TimeSpan ScheduledArrivalTime { get; set; }
        public TimeSpan ScheduledDepartureTime { get; set; }
        public int pudoType { get; set; }
        public string Platform { get; set; }
        public string Line { get; set; }
        public BPLAN_LOC LOC { get; set; }
        public BPLAN_PLT PLT { get; set; }
    }

    // Classes to hold the GTFS output
    // A LIST OF THESE CALENDAR OBJECTS CREATE THE GTFS calendar.txt file
    public class Calendar
    {
        public string service_id { get; set; }
        public int monday { get; set; }
        public int tuesday { get; set; }
        public int wednesday { get; set; }
        public int thursday { get; set; }
        public int friday { get; set; }
        public int saturday { get; set; }
        public int sunday { get; set; }
        public string start_date { get; set; }
        public string end_date { get; set; }
    }

    // A LIST OF THESE TRIPS CREATES THE GTFS trips.txt file.
    public class Trip
    {
        public string route_id { get; set; }
        public string service_id { get; set; }
        public string trip_id { get; set; }
        public string trip_headsign { get; set; }
        public string direction_id { get; set; }
        public string block_id { get; set; }
        public string shape_id { get; set; }
    }

    // A LIST OF THESE STOPTIMES CREATES THE GTFS stop_times.txt file
    public class StopTime
    {
        public string trip_id { get; set; }
        public string arrival_time { get; set; }
        public string departure_time { get; set; }
        public int stop_id { get; set; }
        public int stop_sequence { get; set; }
        public int pickup_type { get; set; }
        public int drop_off_type { get; set; }
    }

    //A LIST OF THESE ATTSTOPS CREATES THE GTFS stops.txt file
    public class GTFSattStop
    {
        public int stop_id { get; set; }
        public string stop_code { get; set; }
        public string stop_name { get; set; }
        public int location_type { get; set; }
        public int parent_station { get; set; }
    }

    // A LIST OF THESE ROUTES CREATES THE GTFS routes.txt file.
    public class Route
    {
        public string route_id { get; set; }
        public string agency_id { get; set; }
        public string route_short_name { get; set; }
        public string route_long_name { get; set; }
        public string route_desc { get; set; }
        public string route_type { get; set; }
        public string route_url { get; set; }
        public string route_color { get; set; }
        public string route_text_color { get; set; }
    }

    // A LIST OF THESE AGENCIES CREATES THE GTFS agencies.txt file.
    public class Agency
    {
        public string agency_id { get; set; }
        public string agency_name { get; set; }
        public string agency_url { get; set; }
        public string agency_timezone { get; set; }
    }
    public class BPLAN_LOC
    {
        public string TIPLOC_x { get; set; }
        public string LocationName_x { get; set; }
        public string StartDate { get; set; }
        public string Easting { get; set; }
        public string Northing { get; set; }
        public string TimingPointType { get; set; }
        public int ZoneResponsible { get; set; }
        public int STANOX { get; set; }
        public string OffNetwork { get; set; }
        public int Quality { get; set; }
        public int index { get; set; }
        public string TIPLOC_y { get; set; }
        public string LocationName_y { get; set; }
    }
    public class BPLAN_PLT
    {
        public string TIPLOC_PlatformID { get; set; }
        public string TIPLOC { get; set; }
        public string PlatformID { get; set; }
        public string StartDate { get; set; }
        public int PlatformLength { get; set; }
        public string PassengerDOO { get; set; }
        public string NonPassengerDOO { get; set; }
        public int index_TIPLOC { get; set; }
        public int index_PlatformID { get; set; }
        public int index { get; set; }
        public string Easting { get; set; }
        public string Northing { get; set; }
        public int Quality { get; set; }
    }
}
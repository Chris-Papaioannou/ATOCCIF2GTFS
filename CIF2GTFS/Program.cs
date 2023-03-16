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

            Console.WriteLine("Clearing temp directory...");
            if (Directory.Exists("temp") == true)
            {
                Directory.Delete("temp", true);
            }
            Directory.CreateDirectory("temp");

            Console.WriteLine("Loading ATT stops");
            List<attStop> attStops = new List<attStop>();
            using (TextReader textReader = File.OpenText("input/cif_tiplocs.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader, CultureInfo.InvariantCulture);
                csvReader.Configuration.Delimiter = ",";
                attStops = csvReader.GetRecords<attStop>().ToList();
            }

            Console.WriteLine("Creating GTFS_STOP_ID keyed dictionary of ATT stops.");
            Dictionary<string, attStop> ATTStopsDictionary = new Dictionary<string, attStop>();
            List<GTFSattStop> GTFSStopsList = new List<GTFSattStop>();
            foreach (attStop attStop in attStops)
            {
                if (attStop.CRS != "" & !ATTStopsDictionary.ContainsKey(attStop.Tiploc))
                {
                    ATTStopsDictionary.Add(attStop.Tiploc, attStop);

                    GTFSattStop gTFSattStop = new GTFSattStop()
                    {
                        stop_id = attStop.Tiploc,
                        stop_code = attStop.CRS,
                        stop_name = attStop.Description,
                        location_type = 1
                    };
                    GTFSStopsList.Add(gTFSattStop);
                }
            }

            Console.WriteLine("Reading the timetable file.");
            List<string> TimetableFileLines = new List<string>(File.ReadAllLines("input/May22.CIF"));
            Dictionary<string, List<StationStop>> StopTimesForJourneyIDDictionary = new Dictionary<string, List<StationStop>>();
            Dictionary<string, JourneyDetail> JourneyDetailsForJourneyIDDictionary = new Dictionary<string, JourneyDetail>();
            string CurrentJourneyID = "";
            string CurrentOperatorCode = "";
            string CurrentTrainType = "";
            string CurrentTrainClass = "";
            string CurrentTrainMaxSpeed = "";
            Calendar CurrentCalendar = null;
            foreach (string TimetableLine in TimetableFileLines)
            {

                if (TimetableLine.StartsWith("BS"))
                {
                    
                    // THIS IS ALMOST CERTAINLY NOT THE CURRENT JOURNEY ID. BUT IT'S OKAY DURING DEVELOPMENT.
                    //CurrentJourneyID = TimetableLine;

                    // Example line is "BSNY244881905191912080000001 POO2D67    111821020 EMU333 100      S            P"
                    CurrentJourneyID = TimetableLine.Substring(2, 7);
                    string StartDateString = TimetableLine.Substring(9, 6);
                    string EndDateString = TimetableLine.Substring(15, 6);
                    string DaysOfOperationString = TimetableLine.Substring(21, 7);

                    // Since a single timetable can have a single Journey ID that is valid at different non-overlapping times a unique Journey ID includes the Date strings and the character at position 79.
                    // CP Note. I'm testing ignoring this line to see if doing so causes any unintended issues or not:
                    // CurrentJourneyID = CurrentJourneyID + StartDateString + EndDateString + TimetableLine.Substring(79, 1);
                    CurrentCalendar = new Calendar()
                    {
                        start_date = "20" + StartDateString,
                        end_date = "20" + EndDateString,
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
                        Location = TimetableLine.Substring(2, 8).Trim()
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

                    if (ATTStopsDictionary.ContainsKey(stationStop.Location))
                    {
                        stationStop.ATTStop = ATTStopsDictionary[stationStop.Location];

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

                }

            }
            
            Console.WriteLine($"Read {StopTimesForJourneyIDDictionary.Keys.Count} journeys.");

            Console.WriteLine("Creating GTFS output.");

            // We have two dictionaries that let us create all our GTFS output.
            // StopTimesForJourneyIDDictionary
            // JourneyDetailsForJourneyIDDictionary
            // CIFStations
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
                    StopTime stopTime = new StopTime()
                    {
                        trip_id = JourneyID + "_trip",
                        stop_id = stationStop.ATTStop.CRS + "_" + stationStop.Platform,
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
            if (Directory.Exists("output_GTFS") == false)
            {
                Directory.CreateDirectory("output_GTFS");
            }
            
            Console.WriteLine("Writing agency.txt");
            TextWriter agencyTextWriter = File.CreateText(@"output_GTFS/agency.txt");
            CsvWriter agencyCSVwriter = new CsvWriter(agencyTextWriter, CultureInfo.InvariantCulture);
            agencyCSVwriter.WriteRecords(AgencyList);
            agencyTextWriter.Dispose();
            agencyCSVwriter.Dispose();

            Console.WriteLine("Writing routes.txt");
            TextWriter routesTextWriter = File.CreateText(@"output_GTFS/routes.txt");
            CsvWriter routesCSVwriter = new CsvWriter(routesTextWriter, CultureInfo.InvariantCulture);
            routesCSVwriter.WriteRecords(RoutesList);
            routesTextWriter.Dispose();
            routesCSVwriter.Dispose();

            Console.WriteLine("Writing trips.txt");
            TextWriter tripsTextWriter = File.CreateText(@"output_GTFS/trips.txt");
            CsvWriter tripsCSVwriter = new CsvWriter(tripsTextWriter, CultureInfo.InvariantCulture);
            tripsCSVwriter.WriteRecords(tripList);
            tripsTextWriter.Dispose();
            tripsCSVwriter.Dispose();

            Console.WriteLine("Writing calendar.txt");
            TextWriter calendarTextWriter = File.CreateText(@"output_GTFS/calendar.txt");
            CsvWriter calendarCSVwriter = new CsvWriter(calendarTextWriter, CultureInfo.InvariantCulture);
            calendarCSVwriter.WriteRecords(calendarList);
            calendarTextWriter.Dispose();
            calendarCSVwriter.Dispose();

            Console.WriteLine("Writing stop_times.txt");
            TextWriter stopTimeTextWriter = File.CreateText("temp/stop_times_full.txt");
            CsvWriter stopTimeCSVwriter = new CsvWriter(stopTimeTextWriter, CultureInfo.InvariantCulture);
            stopTimeCSVwriter.WriteRecords(stopTimesList);
            stopTimeTextWriter.Dispose();
            stopTimeCSVwriter.Dispose();

            Console.WriteLine("Dropping trip IDs with only one matched stop from stop_times.txt");
            ExecProcess("drop_single_stop_trips.py");

            Console.WriteLine("Creating a GTFS .zip file.");
            if (File.Exists("output_GTFS.zip"))
            {
                File.Delete("output_GTFS.zip");
            }
            ZipFile.CreateFromDirectory("output_GTFS", "output_GTFS.zip", CompressionLevel.Optimal, false, Encoding.UTF8);

            Console.WriteLine("You may wish to validate the GTFS output using a tool such as https://github.com/google/transitfeed/");
            ExecProcess("import_GTFS.py");
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
            Console.WriteLine();
            Console.WriteLine("Results:");
            Console.WriteLine(results);

        }

        static TimeSpan stringToTimeSpan(string input)
        {
            // input is expected to be HHMMX where if X = "H" it represents a half-minute
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
        public attStop ATTStop { get; set; }
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
        public string stop_id { get; set; }
        public int stop_sequence { get; set; }
        public int pickup_type { get; set; }
        public int drop_off_type { get; set; }
    }

    //A LIST OF THESE ATTSTOPS CREATES THE GTFS stops.txt file
    public class GTFSattStop
    {
        public string stop_id { get; set; }
        public string stop_code { get; set; }
        public string stop_name { get; set; }
        public int location_type { get; set; }
        public string parent_station { get; set; }
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
    public class attStop
    {
        public string CRS { get; set; }
        public string Tiploc { get; set; }
        public string Description { get; set; }
        public int Stannox { get; set; }
    }
}
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
            Console.WriteLine("Execute python process...");
            Console.WriteLine("Loading ATT stops.");
            if (Directory.Exists("temp") == true)
            {
                Directory.Delete("temp", true);
            }
            Directory.CreateDirectory("temp");

            ExecProcess("merge_cif_stops.py");
            List<attStop> attStops = new List<attStop>();
            using (TextReader textReader = File.OpenText("temp/cif_tiplocs_loc.csv"))
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
                ATTStopsDictionary.Add(attStop.Tiploc, attStop);

                GTFSattStop gTFSattStop = new GTFSattStop()
                {
                    stop_id = attStop.Tiploc,
                    stop_code = attStop.CRS,
                    stop_name = attStop.Description,
                    stop_lat = Math.Round(attStop.YCOORD,5),
                    stop_lon = Math.Round(attStop.XCOORD,5),
                    location_type = 1
                };

                GTFSStopsList.Add(gTFSattStop);
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
                    CurrentJourneyID = CurrentJourneyID + StartDateString + EndDateString + TimetableLine.Substring(79, 1);

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
                    string firstSlot = TimetableLine.Substring(0, 2).Trim();
                    string secondSlot = TimetableLine.Substring(2, 7).Trim();
                    string thirdSlot = TimetableLine.Substring(10, 4).Trim();
                    string fourthSlot = TimetableLine.Substring(15, 4).Trim();
                    string fifthSlot = TimetableLine.Substring(19, 3).Trim();
                    string sixthSlot = TimetableLine.Substring(25, 8).Trim();
                    string seventhSlot = TimetableLine.Substring(33, 3).Trim();

                    if (sixthSlot != "00000000")
                    {
                        StationStop stationStop = new StationStop()
                        {
                            StopType = firstSlot,
                            StationLongCode = secondSlot
                        };
                        if (TimetableLine.StartsWith("LI"))
                        {
                            stationStop.Platform = seventhSlot;
                        }
                        else
                        {
                            stationStop.Platform = fifthSlot;
                        }

                        if (ATTStopsDictionary.ContainsKey(stationStop.StationLongCode))
                        {
                            if (thirdSlot.Count() == 4 && fourthSlot.Count() == 4)
                            {
                                stationStop.WorkingTimetableDepartureTime = stringToTimeSpan(thirdSlot);
                                stationStop.PublicTimetableDepartureTime = stringToTimeSpan(fourthSlot);
                                stationStop.ATTStop = ATTStopsDictionary[stationStop.StationLongCode];

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

                bool JourneyStartedYesterdayFlag = false;
                TimeSpan PreviousStopDepartureTime = new TimeSpan(0);

                foreach (StationStop stationStop in StationStops)
                {
                    if (stationStop.PublicTimetableDepartureTime < PreviousStopDepartureTime)
                    {
                        JourneyStartedYesterdayFlag = true;
                    }

                    StopTime stopTime = new StopTime()
                    {
                        trip_id = JourneyID + "_trip",
                        stop_id = stationStop.ATTStop.Tiploc + "_" + stationStop.Platform,
                        stop_sequence = count
                    };

                    if (JourneyStartedYesterdayFlag == true)
                    {
                        stationStop.WorkingTimetableDepartureTime = stationStop.WorkingTimetableDepartureTime.Add(new TimeSpan(24, 0, 0));
                        stationStop.PublicTimetableDepartureTime = stationStop.PublicTimetableDepartureTime.Add(new TimeSpan(24, 0, 0));
                        stopTime.arrival_time = Math.Floor(stationStop.PublicTimetableDepartureTime.TotalHours).ToString() + stationStop.PublicTimetableDepartureTime.ToString(@"hh\:mm\:ss").Substring(2,6);
                        stopTime.departure_time = Math.Floor(stationStop.PublicTimetableDepartureTime.TotalHours).ToString() + stationStop.PublicTimetableDepartureTime.ToString(@"hh\:mm\:ss").Substring(2, 6);
                    }
                    else
                    {
                        stopTime.arrival_time = stationStop.PublicTimetableDepartureTime.ToString(@"hh\:mm\:ss");
                        stopTime.departure_time = stationStop.PublicTimetableDepartureTime.ToString(@"hh\:mm\:ss");
                    }
                    stopTimesList.Add(stopTime);

                    PreviousStopDepartureTime = stationStop.PublicTimetableDepartureTime;
                    count++;
                }
            }

            List<Calendar> calendarList = JourneyDetailsForJourneyIDDictionary.Values.Select(x => x.OperationsCalendar).ToList();

            Console.WriteLine("Writing agency.txt");
            // write GTFS txts.
            // agency.txt, calendar.txt, calendar_dates.txt, routes.txt, stop_times.txt, stops.txt, trips.txt
            if (Directory.Exists("output") == false)
            {
                Directory.CreateDirectory("output");
            }

            TextWriter agencyTextWriter = File.CreateText(@"output/agency.txt");
            CsvWriter agencyCSVwriter = new CsvWriter(agencyTextWriter, CultureInfo.InvariantCulture);
            agencyCSVwriter.WriteRecords(AgencyList);
            agencyTextWriter.Dispose();
            agencyCSVwriter.Dispose();

            Console.WriteLine("Writing stops.txt");
            TextWriter stopsTextWriter = File.CreateText(@"temp/stations.txt");
            CsvWriter stopsCSVwriter = new CsvWriter(stopsTextWriter, CultureInfo.InvariantCulture);
            stopsCSVwriter.WriteRecords(GTFSStopsList);
            stopsTextWriter.Dispose();
            stopsCSVwriter.Dispose();

            Console.WriteLine("Writing routes.txt");
            TextWriter routesTextWriter = File.CreateText(@"output/routes.txt");
            CsvWriter routesCSVwriter = new CsvWriter(routesTextWriter, CultureInfo.InvariantCulture);
            routesCSVwriter.WriteRecords(RoutesList);
            routesTextWriter.Dispose();
            routesCSVwriter.Dispose();

            Console.WriteLine("Writing trips.txt");
            TextWriter tripsTextWriter = File.CreateText(@"output/trips.txt");
            CsvWriter tripsCSVwriter = new CsvWriter(tripsTextWriter, CultureInfo.InvariantCulture);
            tripsCSVwriter.WriteRecords(tripList);
            tripsTextWriter.Dispose();
            tripsCSVwriter.Dispose();

            Console.WriteLine("Writing calendar.txt");
            TextWriter calendarTextWriter = File.CreateText(@"output/calendar.txt");
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
            if (File.Exists("output.zip"))
            {
                File.Delete("output.zip");
            }
            ZipFile.CreateFromDirectory("output", "output.zip", CompressionLevel.Optimal, false, Encoding.UTF8);

            Console.WriteLine("You may wish to validate the GTFS output using a tool such as https://github.com/google/transitfeed/");
            ExecProcess("import_GTFS.py");
        }

        static void ExecProcess(string my_script)
        {
            // 1) Create Process Info
            var psi = new ProcessStartInfo();
            psi.FileName = @"C:\Program Files\PTV Vision\PTV Visum 2022\Exe\PythonModules\Scripts\python.exe";

            // 2) Provide script and arguments
            var script = my_script;
            
            var ver_path = @"C:\Users\c.papaioannou\PTV Group\TEAM PTV UK - Network Model Phase 3 - General\07 Model Files\01 Detailed Network\DetailedNetwork.ver";

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
            // input is expected to be HHMM
            string hours = input.Substring(0, 2);
            string minutes = input.Substring(2, 2);
            if (hours.StartsWith("0"))
            {
                hours = hours.Substring(1, 1);
            }
            int hoursint = int.Parse(hours);
            int minutesint = int.Parse(minutes);
            TimeSpan timeSpan = new TimeSpan(hoursint, minutesint, 0);
            return timeSpan;
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
        public string StationLongCode { get; set; }
        public string StopType { get; set; } // Origin, Intermediate, or Terminus
        public string Platform { get; set; }
        public TimeSpan WorkingTimetableDepartureTime {get; set;}
        public TimeSpan PublicTimetableDepartureTime { get; set; }
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
        public string stop_headsign { get; set; }
        public string pickup_type { get; set; }
        public string drop_off_type { get; set; }
        public string shape_dist_traveled { get; set; }
    }

    //A LIST OF THESE ATTSTOPS CREATES THE GTFS stops.txt file
    public class GTFSattStop
    {
        public string stop_id { get; set; }
        public string stop_code { get; set; }
        public string stop_name { get; set; }
        public double stop_lat { get; set; }
        public double stop_lon { get; set; }
        public int location_type { get; set; }
        public string parent_station { get; set; }
        //public string vehicle_type { get; set; }
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
        public double XCOORD { get; set; }
        public double YCOORD { get; set; }
    }
}
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Linq;

//////////////////////////////////////////
//
//  Program.cs
//  KronoX Converter by Filip Tripkovic
//  Last updated 2022-08-18
//
//////////////////////////////////////////

namespace KronoXConverter
{
    class Program
    {
        public static readonly ConsoleColor launchExit = ConsoleColor.Cyan;
        public static readonly ConsoleColor primary    = ConsoleColor.White;
        public static readonly ConsoleColor secondary  = ConsoleColor.DarkGray;
        public static readonly ConsoleColor option     = ConsoleColor.Yellow;
        public static readonly ConsoleColor userInput  = ConsoleColor.Green;
        public static readonly ConsoleColor error      = ConsoleColor.Red;

        public static string resourcesFolder { get; private set; }

        private static string downloadsFolder, desktopFolder;
        private static string excelFilePath;
        private static string input;
        private static string recalculation;
        private static bool fileFound, retry, firstFailedSearch;
        private static List<Event> events;
        private static List<string> calendarFilePaths;
        private static List<string> themeFilePaths;
        private static List<string> selectedThemeFilePaths;

        private static char[] charArray;
        private static Stopwatch writeStopwatch;

        public static void Main(string[] args)
        {
#if !DEBUG
            try
#endif
            {
                Launch();

                Search:
                retry = false;
                SearchForCalendarFiles();
                NoCalendarFileFound();
                if (retry) goto Search;
                EnterCalendarFilePathManually();
                if (retry) goto Search;

                SelectTheme();
                SelectRecalculation();
                ResolveAlreadyExistingFile();

                ReadCalendarFiles();
                IncludeCourses();
                ConstructExcelFile();
                //Process.Start(excelFilePath.ToDirectory());
            }
#if !DEBUG
            catch (Exception e)
            {
                WriteLine(error, "Unexpected error");
                WriteLine(secondary, e.ToString());
            }
#endif
            Exit();
        }

        // Step 1
        private static void Launch()
        {
            WriteLine(launchExit, "\nKronoX Converter");
            WriteLine(secondary, "By Filip Tripkovic");
            WriteLine(secondary, "A free open-source console application created with the purpose of re-structuring and combining KronoX schedules");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            resourcesFolder = Environment.CurrentDirectory;
            resourcesFolder = resourcesFolder.Remove(resourcesFolder.LastIndexOf("KronoXConverter") + 15);
            resourcesFolder += "/Resources";

            desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            downloadsFolder = desktopFolder.Replace("Desktop", "Downloads");

            events = new List<Event>();
            calendarFilePaths = new List<string>();
            selectedThemeFilePaths = new List<string>();
            firstFailedSearch = true;
        }

        // Step 2
        private static void SearchForCalendarFiles()
        {
            fileFound = false;
            SearchFolder(downloadsFolder);
            if (input != "y")
                SearchFolder(desktopFolder);

            void SearchFolder(string folder)
            {
                foreach (string file in Directory.GetFiles(folder))
                {
                    if (file.EndsWith(".ics"))
                    {
                        fileFound = true;
                        WriteLine(primary, "\nCalendar file found: " + file.Substring(file.LastIndexOf("/") + 1));
                        WriteLine(secondary, "At: " + file);
                        Write(primary, "Do you want to use this file" + (calendarFilePaths.Count == 0 ? "?" : " as well?"));
                        WriteLine(option, " y/n");

                        UserInput:
                        Console.ForegroundColor = userInput;
                        input = Console.ReadKey().KeyChar.ToString().ToLower();
                        Console.WriteLine();

                        if (input == "y")
                        {
                            calendarFilePaths.Add(file);
                        }

                        else if (input != "n")
                        {
                            WriteLine(error, "Invalid input");
                            goto UserInput;
                        }

                        // If "n": Continue to next file
                    }
                }
            }
        }

        // Step 2.1
        private static void NoCalendarFileFound()
        {
            if (fileFound)
                // Continue to EnterCalendarFilePathManually()
                return;

            Write(primary, "\nNo calendar files found in downloads or desktop");

            if (!firstFailedSearch)
                // Continue to EnterCalendarFilePathManually()
                return;

            firstFailedSearch = false;
            Write(primary, "\nHave you downloaded a calendar file from KronoX's website yet?");
            WriteLine(option, " y/n ");

            UserInput:
            Console.ForegroundColor = userInput;
            input = Console.ReadKey().KeyChar.ToString();
            Console.WriteLine();

            if (input == "n")
            {
                Write(primary, "\nYou must download one or more calendar files from KronoX and place them" +
                   " in your downloads folder or on your desktop. Continue by locating your schedules," +
                   " download them by clicking 'Get iCal file' / 'Hämta iCal fil', then enter");
                Write(option, " retry ");
                WriteLine(primary, "to do another search");

                UserInput2:
                Console.ForegroundColor = userInput;
                input = Console.ReadLine().ToString().ToLower();

                if (input != "retry")
                {
                    WriteLine(error, "Invalid input");
                    goto UserInput2;
                }

                // If "retry": retry serach process
                retry = true;
            }

            else if (input != "y")
            {
                WriteLine(error, "Invalid input");
                goto UserInput;
            }

            // If "y": Continue to EnterCalendarFilePathManually()
        }

        // Step 2.2
        private static void EnterCalendarFilePathManually()
        {
            if (calendarFilePaths.Count != 0)
                return;

            if (fileFound)
                Write(primary, "\nNo more calendar files found in downloads or desktop");

            Write(primary, "\nEnter path for calender file manually, or move the file(s) to downloads or desktop, then enter");
            Write(option, " retry ");
            WriteLine(primary, "to do another search");

            UserInput:
            Console.ForegroundColor = userInput;
            input = Console.ReadLine().ToString();

            if (input.ToLower() == "retry")
            {
                retry = true;
            }

            else if (File.Exists(input))
            {
                if (!input.EndsWith(".ics"))
                {
                    WriteLine(error, "File must end with .ics");
                    goto UserInput;
                }

                calendarFilePaths.Add(input);
            }

            else if (File.Exists(input + ".ics"))
            {
                calendarFilePaths.Add(input + ".ics");
            }

            else
            {
                WriteLine(error, "File not found");
                goto UserInput;
            }
        }

        // Step 3
        private static void SelectTheme()
        {
            Write(primary, "\nSelect a theme by entering a number, or enter");
            Write(option, " all ");
            WriteLine(primary, "to try all themes and decide later");

            themeFilePaths = Directory.GetFiles(resourcesFolder + "/Themes").Where(item => item.EndsWith(".txt")).ToList();
            themeFilePaths.Sort();
            for (int i = 0; i < themeFilePaths.Count; ++i)
            {
                Write(option, (i + 1) + " ");
                WriteLine(primary, themeFilePaths[i].ToFileName(false));
            }

            UserInput:
            Console.ForegroundColor = userInput;
            input = Console.ReadLine().ToString();

            if (input == "all")
            {
                selectedThemeFilePaths = themeFilePaths;
                return;
            }

            try { selectedThemeFilePaths.Add(themeFilePaths[int.Parse(input) - 1]); }
            catch
            {
                WriteLine(error, "Invalid input");
                goto UserInput;
            }
        }

        // Step 4
        private static void SelectRecalculation()
        {
            Write(primary, "\nSelect recalculation interval");
            WriteLine(option, " 1/2/3");
            WriteLine(secondary, "Each event and day in your spreadsheet schedule will automatically be grayed" +
                " out once it has ended. This can only happen whenever the spreadsheet is updated. Your spreadsheet" +
                " can be updated automatically, this is in Google Sheets called recalculation interval, and can be" +
                " changed later by opening your spreadshet and going to File > Settings > Calculation");
            
            Write(option, "\n1 ");
            Write(primary, "On change and every hour");
            WriteLine(secondary, " (better performance)");
            WriteLine(secondary, "If you plan on having your spreadsheet open at all times, and your schedule is rather large");

            Write(option, "\n2 ");
            Write(primary, "On change and every minute");
            WriteLine(secondary, " (better accuary)");
            WriteLine(secondary, "If you plan on having your spreadsheet open at all times, and your schedule is rather small");

            Write(option, "\n3 ");
            Write(primary, "On change");
            WriteLine(secondary, " (no automatic updating)");
            WriteLine(secondary, "If you plan on opening and closing it frequently");

            Write(primary, "\nIf unsure of which to use,");
            Write(option, " 1 ");
            WriteLine(primary, "is recommended");
            WriteLine(secondary, "This setting can be changed later by opening your spreadsheet and going to File > Settings > Calculation");

            UserInput:
            Console.ForegroundColor = userInput;
            recalculation = Console.ReadKey().KeyChar.ToString();
            Console.WriteLine();

            if (recalculation != "1" && recalculation != "2" && recalculation != "3")
            {
                WriteLine(error, "Invalid input");
                goto UserInput;
            }

            // If "1" or "2" or "3": Continue to ResolveAlreadyExistingFile()
        }

        // Step 5
        private static void ResolveAlreadyExistingFile()
        {
            excelFilePath = calendarFilePaths[0].ToDirectory() + "/Setup Sheet 1.xlsx";

            if (!File.Exists(excelFilePath))
                return;

            WriteLine(primary, "\nThe file to be created: " + excelFilePath.ToFileName(true) + ", already exists");
            WriteLine(secondary, "At: " + excelFilePath);
            Write(primary, "Do you want to replace it?");
            WriteLine(option, " y/n");

            UserInput:
            Console.ForegroundColor = userInput;
            input = Console.ReadKey().KeyChar.ToString().ToLower();
            Console.WriteLine();

            if (input == "y")
                File.Delete(excelFilePath);

            else if (input == "n")
                excelFilePath = excelFilePath.NextUniqeName();

            else
            {
                WriteLine(error, "Invalid input");
                goto UserInput;
            }
        }

        // Step 6
        private static void ReadCalendarFiles()
        {
            foreach (string calendarFilePath in calendarFilePaths)
                CalendarReader.ReadCalendarFile(events, calendarFilePath);

            events = events.Distinct().ToList();
            events.Sort((ev1, ev2) => DateTime.Compare(ev1.start, ev2.start));
        }

        // Step 7
        private static void IncludeCourses()
        {
            List<string> courses = events.Select(ev => ev.course).Distinct().ToList();
            courses.Sort();

            WriteLine(primary, "\nCourses found: " + courses.Count);
            WriteLine(primary, "Select whether or not to include the following courses:");

            foreach (string course in courses)
            {
                List<Event> eventsInCourse = events.Where(ev => ev.course == course).ToList();

                Write(primary, " - " + course);
                Write(secondary, " (" + eventsInCourse.Count + " events)");
                Write(option, " y/n ");

                if (eventsInCourse.Count < 4)
                {
                    int left = Console.CursorLeft;
                    int top = Console.CursorTop;

                    foreach (Event ev in eventsInCourse)
                        Write(secondary, "\n" + ev.start.ToString("   yyyy-MM-dd ") + ev.description);

                    Console.SetCursorPosition(left, top - eventsInCourse.Count);
                }

                UserInput:
                Console.ForegroundColor = userInput;
                input = Console.ReadKey().KeyChar.ToString().ToLower();
                Console.WriteLine();

                if (input == "n")
                    events.RemoveAll(ev => ev.course == course);

                else if (input != "y")
                {
                    WriteLine(error, "Invalid input");
                    goto UserInput;
                }

                if (eventsInCourse.Count < 4)
                {
                    Console.SetCursorPosition(Console.WindowWidth - 1, Console.WindowHeight - 1);
                    Console.WriteLine();
                }
            }
        }

        // Step 8
        private static void ConstructExcelFile()
        {
            WriteLine(primary, "\nConstructing " + excelFilePath.ToFileName(true) + " ...");
            WriteLine(secondary, "At: " + excelFilePath);
            Stopwatch stopwatchTheme = Stopwatch.StartNew();
            Stopwatch stopwatchFile = Stopwatch.StartNew();
            ExcelFormatter.NewSetupPackage(excelFilePath, recalculation, events);

            foreach (string selectedThemeFilePath in selectedThemeFilePaths)
            {
                Theme.Initialize(selectedThemeFilePath);
                if (!Theme.successful)
                {
                    WriteLine(primary, "Skipping theme " + selectedThemeFilePath.ToFileName(false));
                    continue;
                }

                stopwatchTheme.Restart();
                ExcelFormatter.AddScheduleSheet(events);
                stopwatchTheme.Stop();
                WriteLine(secondary, " (" + stopwatchTheme.Elapsed.TotalSeconds.ToString("0.0") + " sec)");
            }

            ExcelFormatter.SavePackage();
            stopwatchFile.Stop();
            Write(primary, "Finished!");
            WriteLine(secondary, " (" + stopwatchFile.Elapsed.TotalSeconds.ToString("0") + " sec)");
            WriteLine(primary, "\nNext: Drag " + excelFilePath.ToFileName(true) +
                " into your Google Drive, open it, and follow the instructions to complete the setup");
        }

        // Step 9
        private static void Exit()
        {
            WriteLine(launchExit, "\nKronoX Converter closing");
            Console.ResetColor();
            Environment.Exit(0);
        }

        /// <summary>
        /// Writes the provided text to the console in the provided color character for character
        /// </summary>
        public static void Write(ConsoleColor consoleColor, string text, bool instant = false)
        {
            Console.ForegroundColor = consoleColor;

            if (instant)
            {
                Console.Write(text);
                return;
            } 

            charArray = text.ToCharArray();
            writeStopwatch = Stopwatch.StartNew();

            for (int i = 0; i < charArray.Length; ++i)
            {
                // Syncronous wait
                while (writeStopwatch.ElapsedMilliseconds < 1) ;
                Console.Write(charArray[i]);
                writeStopwatch.Restart();
            }

            writeStopwatch.Reset();
        }

        /// <summary>
        /// Writes the provided text to the console in the provided color character for character
        /// </summary>
        public static void WriteLine(ConsoleColor consoleColor, string text, bool instant = false)
        {
            Write(consoleColor, text + "\n", instant);
        }

        /// <summary>
        /// Clears the current line in the console
        /// </summary>
        public static void ClearLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }
    }
}

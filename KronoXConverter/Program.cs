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
//  Last updated 2022-03-13
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
        private static string calendarFilePath, excelFilePath;
        private static string input;
        private static string recalculation;
        private static bool fileFound, retry, firstFailedSearch;
        private static List<Event> events;
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
                int a = int.Parse("a");
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

                ReadCalendarFile(events);
                ConstructExcelFile(events);
                //Process.Start(excelFilePath.ToDirectory());
            }
#if !DEBUG
            catch (Exception e)
            {
                WriteLine(error, "Unexpected fatal error");
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
            WriteLine(primary, "\nA free open-source console application created with the specific purpose of re-structuring KronoX schedules");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            resourcesFolder = Environment.CurrentDirectory;
            resourcesFolder = resourcesFolder.Remove(resourcesFolder.LastIndexOf("KronoXConverter") + 15);
            resourcesFolder += "/Resources";

            downloadsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/Downloads";
            desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            events = new List<Event>();
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
                        ClearLine();
                        Write(primary, "Do you want to use this file?");
                        WriteLine(option, " y/n");

                        UserInput:
                        Console.ForegroundColor = userInput;
                        input = Console.ReadKey().KeyChar.ToString().ToLower();
                        Console.WriteLine();

                        if (input == "y")
                        {
                            calendarFilePath = file;
                            break;
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

        // Step 3
        private static void NoCalendarFileFound()
        {
            if (fileFound)
                // Continue to EnterCalendarFilePathManually()
                return;

            Write(primary, "\nNo calendar file found in downloads or desktop");

            if (!firstFailedSearch)
                // Continue to EnterCalendarFilePathManually()
                return;

            firstFailedSearch = false;
            Write(primary, "\nHave you downloaded a calendar file from KronoX website yet?");
            WriteLine(option, " y/n ");

            UserInput:
            Console.ForegroundColor = userInput;
            input = Console.ReadKey().KeyChar.ToString();
            Console.WriteLine();

            if (input == "n")
            {
                Write(primary, "\nYou must download a calendar file from KronoX and place it" +
                   " in your downloads folder or on your desktop. Continue by locating your schedule," +
                   " download it by clicking 'Get iCal file' / 'Hämta iCal fil', then enter");
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

        // Step 4
        private static void EnterCalendarFilePathManually()
        {
            if (calendarFilePath != null)
                return;

            if (fileFound)
                Write(primary, "\nNo more calendar files found in downloads or desktop");

            Write(primary, "\nEnter path for calender file manually, or move the file to downloads or desktop and enter");
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

                calendarFilePath = input;
            }

            else if (File.Exists(input + ".ics"))
            {
                calendarFilePath = input + ".ics";
            }

            else
            {
                WriteLine(error, "File not found");
                goto UserInput;
            }
        }

        // Step 5
        private static void SelectTheme()
        {
            Write(primary, "\nSelect a theme by entering a number, or enter");
            Write(option, " all ");
            WriteLine(primary, "to try all of them and decide later");

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

        // Step 6
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

        // Step 7
        private static void ResolveAlreadyExistingFile()
        {
            excelFilePath = calendarFilePath.ToDirectory() + "/Setup Sheet 1.xlsx";

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

        // Step 8
        private static void ReadCalendarFile(List<Event> events)
        {
            CalendarReader.ReadCalendarFile(events, calendarFilePath);
        }

        // Step 9
        private static void ConstructExcelFile(List<Event> events)
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

        // Exit
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

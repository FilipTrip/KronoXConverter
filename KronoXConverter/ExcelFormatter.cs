using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;

//////////////////////////////////////////
//
//  ExcelFormatter.cs
//  KronoX Converter by Filip Tripkovic
//  Last updated 2022-03-13
//
//////////////////////////////////////////

namespace KronoXConverter
{
    public static class ExcelFormatter
    {
        private static Dictionary<string, string> abbreviationCells; // Key: course, Value: Cell containing an abbreviation
        private static Dictionary<string, int> courseColors; // Key: course, Value: Theme.fillCourse index
        private static List<DayOfWeek> daysOfWeek; // Monday -> Sunday

        private static List<List<Event>> eventWeeks;
        private static int eventCount;
        private static int row, firstScheduleRow;

        private static Dictionary<string, string> formulas;                        // Key: range, Value: formula
        private static Dictionary<string, string> conditionalFormattingsDateEvent; // Key: range, Value: formula
        private static Dictionary<string, string> conditionalFormattingsWeek;      // Key: range, Value: formula

        private static ExcelPackage package;
        private static ExcelRange cells;
        private static ExcelWorksheet worksheet;

        private static IExcelConditionalFormattingExpression conditionalFormatting;


        /// <summary>
        /// Initializes and prepares setup package
        /// </summary>
        public static void NewSetupPackage(string filePath, string recalculation, List<Event> events)
        {
            // Delete file if it already exists
            if (File.Exists(filePath))
            {
                Program.WriteLine(Program.secondary,"File was deleted by ExcelFormatter," +
                    " should not have been neccessary (you can ignore this warning)");
                File.Delete(filePath);
            }

            // Initialize new package
            package = new ExcelPackage(filePath);
            package.Workbook.CalcMode = ExcelCalcMode.Manual;

            // Add setup sheet
            using (ExcelPackage setupPackage = new ExcelPackage(Program.resourcesFolder + "/SpreadSheets/Setup.xlsx"))
            {
                worksheet = package.Workbook.Worksheets.Add("Setup 1", setupPackage.Workbook.Worksheets[0]);
            }

            // Save recalculation setting
            worksheet.Cells["A1"].Value = recalculation;

            // Create days of week list
            daysOfWeek = new List<DayOfWeek>();
            for (int i = 0; i < 7; ++i)
                daysOfWeek.Add((DayOfWeek)Enum.GetValues(typeof(DayOfWeek)).GetValue((i + 1) % 7));

            eventWeeks = new List<List<Event>>();

            int week = -1;
            int weekIndex = -1;

            foreach (Event ev in events)
            {
                // If new week
                if (ev.week != week)
                {
                    // Add new eventWeek and increment to it
                    eventWeeks.Add(new List<Event>());
                    ++weekIndex;
                    week = ev.week;
                }

                //Add event to current eventWeek
                eventWeeks[weekIndex].Add(ev);
            }

            eventCount = events.Count;
        }

        /// <summary>
        /// Creates, format and saves Excel file to the provided file path
        /// </summary>
        public static void AddScheduleSheet(List<Event> events)
        {
            //
            // 1. Add new schedule sheet
            //

            using (ExcelPackage schedulePackage = new ExcelPackage(Program.resourcesFolder + "/SpreadSheets/Schedule.xlsx"))
            {
                worksheet = package.Workbook.Worksheets.Add(Theme.name, schedulePackage.Workbook.Worksheets[0]);
            }

            formulas = new Dictionary<string, string>();
            conditionalFormattingsDateEvent = new Dictionary<string, string>();
            conditionalFormattingsWeek = new Dictionary<string, string>();
            worksheet.Hidden = eWorkSheetHidden.Hidden;
            worksheet.DeleteColumn(14, 20);
            cells = worksheet.Cells;
            ReformatHeader();

            //
            // 2. Add and write abbreviatinos
            //

            abbreviationCells = new Dictionary<string, string>();
            courseColors = new Dictionary<string, int>();
            row = 14; // Row of first abbreviation

            WriteAbbreviation("w", "Week", row++);
            
            foreach (string course in events.Select(item => item.course).Distinct().ToList())
            {
                // Make abbreviation and remove any trailing spaces
                string abbreviation = course.Substring(0, 5);
                while (abbreviation.Length > 0 && abbreviation[abbreviation.Length - 1] == ' ')
                    abbreviation = abbreviation.Remove(abbreviation.Length - 1);

                // Write and add abbreviation, add course color
                WriteAbbreviation(abbreviation, course, row++);
                courseColors.Add(course, courseColors.Count);
            }

            firstScheduleRow = row += 3;

            //
            // 3. Construct schedule
            //

            List<Event> eventWeek;
            Event ev = new Event();
            DateTime oldStart;

            int eventsWritten = 0;
            int dateRow = 0;
            int fillCoursesCount = Theme.fillCourses.Count();

            DayOfWeek expectedDayOfWeek;

            // For each week

            for (int weekIndex = 0; weekIndex < eventWeeks.Count; ++weekIndex)
            {
                eventWeek = eventWeeks[weekIndex];
                expectedDayOfWeek = DayOfWeek.Monday;
                InsertBlankRow(++row);

                // Write week of first event
                Event firstEv = eventWeek[0];
                cells["D" + row].Formula = abbreviationCells["Week"] + " & " + firstEv.week;

                // Write month of monday of this week
                DateTime monday = firstEv.start.AddDays(-DayIndex(firstEv.start.DayOfWeek));
                cells["E" + row].Formula = "PROPER(TEXT(" + monday.Month + "*29; \"MMMM\"))";

                // Add conditional formatting to week
                cells["C" + row].Style.Font.Color.SetColor(Theme.fillBackground);
                AddConditionalFormattingWeek(row, monday.AddDays(6));

                // For each event
                for (int eventIndex = 0; eventIndex < eventWeeks[weekIndex].Count; ++eventIndex)
                {
                    // If expected day of week is prior to event day of week
                    if (DayIndex(eventWeek[eventIndex].start.DayOfWeek) > DayIndex(expectedDayOfWeek))
                    {
                        // Insert weekday row and leave it empty
                        InsertRowWeekday(++row, expectedDayOfWeek);

                        DateTime expectedDate = eventWeek[eventIndex].start.AddDays(DayIndex(expectedDayOfWeek) - DayIndex(eventWeek[0].start.DayOfWeek));
                        cells["D" + row].Formula = "PROPER(TEXT(DATEVALUE(\"" + expectedDate.ToString("yyyy-MM-dd") + "\"); \"DDD\"))";
                        cells["E" + row].Value = expectedDate.Day;

                        AddConditionalFormattingDate(row, expectedDate);

                        // Expect next day of week
                        expectedDayOfWeek = daysOfWeek[(DayIndex(expectedDayOfWeek) + 1) % daysOfWeek.Count];

                        // Attempt same event again
                        --eventIndex;
                        continue;
                    }

                    // Write progress

                    Program.ClearLine();
                    Program.Write(Program.primary, Theme.name + ": " + (70f * ++eventsWritten / eventCount).ToString("0.0") + "%", true);

                    // Next event

                    oldStart = ev.start;
                    ev = eventWeek[eventIndex];
                    InsertRowWeekday(++row, ev.start.DayOfWeek);

                    // If new date
                    if (!SameDate(ev.start, oldStart))
                    {
                        // Write new date and add conditional formatting to it
                        cells["D" + row].Formula = "PROPER(TEXT(DATEVALUE(\"" + ev.start.ToString("yyyy-MM-dd") + "\"); \"DDD\"))";
                        cells["E" + row].Value = ev.start.Day;

                        AddConditionalFormattingDate(row, ev.start);
                        dateRow = row;
                    }

                    // Write event information and add conditional formatting to it

                    cells["F" + row].Value = ev.start.ToString("HH:mm") + " - " + ev.endHour.ToString("00") + ":" + ev.endMin.ToString("00");
                    cells["I" + row].Value = ev.description;
                    cells["J" + row].Value = ev.location;

                    var courseCell = cells["G" + row];
                    courseCell.Formula = abbreviationCells[ev.course];

                    if (Theme.hasCourseColors)
                        courseCell.Style.Fill.BackgroundColor.SetColor(Theme.fillCourses[courseColors[ev.course] % fillCoursesCount]);

                    AddConditionalFormattingEvent(row, dateRow, ev);

                    // If more events exists this week
                    if (eventIndex + 1 < eventWeek.Count)
                    {
                        // And next event is another day
                        if (!SameDate(ev.start, eventWeek[eventIndex + 1].start))
                        {
                            // Expect next day of week
                            expectedDayOfWeek = daysOfWeek[(DayIndex(expectedDayOfWeek) + 1) % daysOfWeek.Count];
                        }
                    }

                    // No more events this week
                    else
                    {
                        // Expect next day of week
                        expectedDayOfWeek = daysOfWeek[(DayIndex(expectedDayOfWeek) + 1) % daysOfWeek.Count];

                        // Add empty weekday rows up to friday
                        while (DayIndex(expectedDayOfWeek) <= DayIndex(DayOfWeek.Friday))
                        {
                            // Create weekday row and leave it empty
                            InsertRowWeekday(++row, expectedDayOfWeek);

                            DateTime dateTime = ev.start.AddDays(DayIndex(expectedDayOfWeek) - DayIndex(ev.start.DayOfWeek));
                            cells["D" + row].Formula = "PROPER(TEXT(DATEVALUE(\"" + dateTime.ToString("yyyy-MM-dd") + "\"); \"DDD\"))";
                            cells["E" + row].Value = dateTime.Day;

                            // Expect next day of week
                            expectedDayOfWeek = daysOfWeek[(DayIndex(expectedDayOfWeek) + 1) % daysOfWeek.Count];
                        }
                    }
                }

                InsertBlankRow(++row);
            }

            //
            // 3. Write stored formulas to cells
            // (Significant increase in performance when adding formulas after, rather than in the event loop)
            //

            int i = 0;
            int formulaCount = formulas.Count() + conditionalFormattingsDateEvent.Count() + conditionalFormattingsWeek.Count();

            void WritePercentage()
            {
                Program.ClearLine();
                Program.Write(Program.primary, Theme.name + ": " + (70 + 30f * ++i / formulaCount).ToString("0.0") + "%", true);
            }

            foreach (KeyValuePair<string, string> formula in formulas)
            {
                WritePercentage();
                foreach (var cell in cells[formula.Key])
                {
                    cell.Formula = formula.Value;
                }
            }

            foreach (KeyValuePair<string, string> formatting in conditionalFormattingsDateEvent)
            {
                WritePercentage();
                foreach (var cell in cells[formatting.Key])
                {
                    conditionalFormatting = cell.ConditionalFormatting.AddExpression();
                    conditionalFormatting.Formula = formatting.Value;
                    conditionalFormatting.Style.Font.Color.SetColor(Theme.fontColorEnded);
                }
            }

            foreach (KeyValuePair<string, string> formatting in conditionalFormattingsWeek)
            {
                WritePercentage();
                foreach (var cell in cells[formatting.Key])
                {
                    conditionalFormatting = cell.ConditionalFormatting.AddExpression();
                    conditionalFormatting.Formula = formatting.Value;
                    conditionalFormatting.Style.Font.Color.SetColor(Theme.fontColorSecondary);
                }
            }

            //
            // 4. Final adjustments to schedule
            //

            // Add 10 blank rows
            for (int max = row + 10; row <= max;)
                InsertBlankRow(++row);

            // Group header
            worksheet.Rows[3, firstScheduleRow - 1].OutlineLevel = 1;
        }

        /// <summary>
        /// Saves and disposes package
        /// </summary>
        public static void SavePackage()
        {
            package.Workbook.Calculate();
            package.Save();
            package.Dispose();
        }

        /// <summary>
        /// Inserts blank row
        /// </summary>
        private static void InsertBlankRow(int row)
        {
            worksheet.InsertRow(row, 1, 2);
        }

        /// <summary>
        /// Inserts and styles row for weekday
        /// </summary>
        private static void InsertRowWeekday(int row, DayOfWeek dayOfWeek)
        {
            // Insert row
            worksheet.InsertRow(row, 1, 2);

            // Style cells
            Color color = DayIndex(dayOfWeek) % 2 == 0 ? Theme.fillScheduleDark : Theme.fillScheduleLight;
            worksheet.Cells["C" + row + ",K" + row].Style.Font.Color.SetColor(color);
            worksheet.Cells["C" + row + ":K" + row].Style.Fill.SetBackground(color);
            worksheet.Cells["D" + row + ":F" + row].Style.Font.Color.SetColor(Theme.fontColorSchedule);
            worksheet.Cells["I" + row + ":J" + row].Style.Font.Color.SetColor(Theme.fontColorSchedule);
        }

        /// <summary>
        /// Inserts a row and writes both the abbreviation and the full form
        /// </summary>
        private static void WriteAbbreviation(string abbreviation, string fullForm, int row)
        {
            InsertBlankRow(row);
            worksheet.Cells["D" + row].Value = abbreviation;
            worksheet.Cells["F" + row].Value = fullForm;
            worksheet.Cells["F" + row].Style.Font.Color.SetColor(Theme.fontColorSecondary);
            abbreviationCells.Add(fullForm, "D" + row);
        }

        /// <summary>
        /// Returns whether or not the two events occur on the same date
        /// </summary>
        private static bool SameDate(DateTime dt1, DateTime dt2)
        {
            return
                dt1.Day   == dt2.Day &&
                dt1.Month == dt2.Month &&
                dt1.Year  == dt2.Year;
        }

        /// <summary>
        /// Returns the index of this DayOfWeek (Monday = 0)
        /// </summary>
        private static int DayIndex(DayOfWeek dayOfWeek)
        {
            return ((int)dayOfWeek + 6) % 7;
        }

        /// <summary>
        /// Adds conditional formatting to an event
        /// </summary>
        private static void AddConditionalFormattingEvent(int row, int dateRow, Event ev)
        {
            // OR
            // (
            //     DATEVALUE("yyy-MM-dd") < DATEVALUE(F1);
            //     AND
            //     (
            //         DATEVALUE("yyy-MM-dd") = DATEVALUE(F1);
            //         TIMEVALUE("HH.mm") < TIMEVALUE(G1)
            //     )
            // )

            string dateHasPassed = "C" + dateRow;
            string sameDate = "DATEVALUE(\"" + ev.start.ToString("yyyy-MM-dd") + "\") = DATEVALUE(F1)";
            string timeHasPassed = "TIMEVALUE(\"" + ev.endHour.ToString("00") + ":" + ev.endMin.ToString("00") + "\") < TIMEVALUE(G1)";

            formulas.Add("K" + row, "OR(" + dateHasPassed + ";AND(" + sameDate + ";" + timeHasPassed + "))");
            conditionalFormattingsDateEvent.Add("F" + row + ",I" + row + ":J" + row, "K" + row);
        }

        /// <summary>
        /// Adds conditional formatting to a date
        /// </summary>
        private static void AddConditionalFormattingDate(int row, DateTime dateTime)
        {
            // DATEVALUE("yyy-MM-dd") < DATEVALUE(F1)

            formulas.Add("C" + row, "DATEVALUE(\"" + dateTime.ToString("yyyy-MM-dd") + "\") < DATEVALUE(F1)");
            conditionalFormattingsDateEvent.Add("D" + row + ":E" + row, "C" + row);
        }

        /// <summary>
        /// Adds conditional formatting to a date
        /// </summary>
        private static void AddConditionalFormattingWeek(int row, DateTime dateTime)
        {
            // DATEVALUE("yyy-MM-dd") < DATEVALUE(F1)

            formulas.Add("C" + row, "DATEVALUE(\"" + dateTime.ToString("yyyy-MM-dd") + "\") < DATEVALUE(F1)");
            conditionalFormattingsWeek.Add("D" + row + ":E" + row, "C" + row);
        }

        /// <summary>
        /// Changes colors and typefaces of the header to match the selected theme
        /// </summary>
        private static void ReformatHeader()
        {
            // Header
            var headerCells = cells["1:18"];
            headerCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerCells.Style.Fill.BackgroundColor.SetColor(Theme.fillBackground);
            headerCells.Style.Font.Color.SetColor(Theme.fontColorPrimary);
            headerCells.Style.Font.Name = Theme.typefacePrimary;

            // Row 1
            worksheet.Cells["1:1"].Style.Fill.BackgroundColor.SetColor(Theme.fillHeader);
            worksheet.Cells["1:1"].Style.Font.Color.SetColor(Theme.fontColorHeader);

            // Button
            ExcelPicture updateButton = worksheet.Drawings[0] as ExcelPicture;
            updateButton.Image = Image.FromFile(Program.resourcesFolder + "/Images/" + Theme.updateButton);

            // Headings fonts
            worksheet.Cells["D5" ].Style.Font.Name = Theme.typefaceHeading;
            worksheet.Cells["D12"].Style.Font.Name = Theme.typefaceHeading;
            worksheet.Cells["D16"].Style.Font.Name = Theme.typefaceHeading;

            // Heading lines
            worksheet.Cells["C7:K7"  ].Style.Border.Bottom.Color.SetColor(Theme.fontColorPrimary);
            worksheet.Cells["C12:K12"].Style.Border.Bottom.Color.SetColor(Theme.fontColorPrimary);
            worksheet.Cells["C16:K16"].Style.Border.Bottom.Color.SetColor(Theme.fontColorPrimary);

            // Text colors
            cells["D6"].Style.Font.Color.SetColor(Theme.fontColorSecondary);
            cells["D9"].Style.Font.Color.SetColor(Theme.fontColorSecondary);
            cells["G2"].Style.Font.Color.SetColor(Theme.fontColorCourse);
        }
    }
}

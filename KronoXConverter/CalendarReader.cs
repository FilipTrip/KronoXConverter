using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

//////////////////////////////////////////
//
//  CalendarReader.cs
//  KronoX Converter by Filip Tripkovic
//  Last updated 2022-07-30
//
//////////////////////////////////////////

namespace KronoXConverter
{
    public static class CalendarReader
    {
        /// <summary>
        /// Token to index
        /// </summary>
        private enum SummaryToken { Course, Sign, Description, Resourse, Programme };

        /// <summary>
        /// Tokens used by KronoX to divide summaries
        /// </summary>
        private static readonly string[] summaryTokensSwe = { "Kurs.grp: ", "Sign: ", "Moment: ", "Hjälpm.: ", "Program: " };

        /// <summary>
        /// Tokens used by KronoX to divide summaries
        /// </summary>
        private static readonly string[] summaryTokensEng = { "Coursegrp: ", "Sign: ", "Description: ", "Resourse: ", "Programme: " };

        /// <summary>
        /// Returns a substring which lies between the provided start and end string
        /// </summary>
        private static string SubstringBetween(string s, string start, string end)
        {
            if (!s.Contains(start) || !s.Contains(end))
                return s;

            int startPos = s.IndexOf(start) + start.Length;
            int endPos = s.IndexOf(end, startPos);

            return s.Substring(startPos, endPos - startPos);
        }

        /// <summary>
        /// Extracts and returns the context of the provided token, if present in the summary
        /// </summary>
        private static string ExtractFromSummary(string summary, SummaryToken summaryToken)
        {
            int i = (int)summaryToken;
            int startPos;
            int tokenPos;
            int endPos = int.MaxValue;

            if (summary.Contains(summaryTokensSwe[i]))
                startPos = summary.IndexOf(summaryTokensSwe[i]) + summaryTokensSwe[i].Length;

            else if (summary.Contains(summaryTokensEng[i]))
                startPos = summary.IndexOf(summaryTokensEng[i]) + summaryTokensEng[i].Length;

            else
                return "";

            foreach (string token in summaryTokensSwe)
            {
                tokenPos = summary.IndexOf(token);
                if (tokenPos > startPos && tokenPos < endPos)
                    endPos = tokenPos;
            }

            foreach (string token in summaryTokensEng)
            {
                tokenPos = summary.IndexOf(token);
                if (tokenPos > startPos && tokenPos < endPos)
                    endPos = tokenPos;
            }

            if (endPos == int.MaxValue)
                return summary.Substring(startPos);

            return summary.Substring(startPos, endPos - startPos);
        }

        /// <summary>
        /// Reads the 
        /// </summary>
        public static void ReadCalendarFile(List<Event> events, string filePath)
        {
            // Open, read and close file

            string[] lines = File.ReadAllLines(filePath);

            // Interpret lines

            string line, nextLine;
            Event ev = new Event();
            Calendar calendar = new GregorianCalendar(GregorianCalendarTypes.Localized);

            for (int i = 0; i < lines.Length; ++i)
            {
                // Get line
                line = lines[i];
        
                // If following line starts with a tab or space, add it to this line
                for (; i + 1 < lines.Length; ++i)
                {
                    nextLine = lines[i + 1];
            
                    // If next line starts with '\t' or ' '
                    // treat it as a continuation of the current line
                    if (nextLine[0] == '\t' || nextLine[0] == ' ')
                        line += nextLine;
                    else
                        break;
                }

                // Calendar Event Structure:

                // BEGIN:VEVENT
                // DTSTART:YYYYMMDDTHHMMSSZ
                // DTEND:YYYYMMDDTHHMMSSZ
                // DTSTAMP:YYYYMMDDTHHMMSSZ
                // UID:___
                // CREATED:YYYYMMDDTHHMMSSZ
                // LAST-MODIFIED:YYYYMMDDTHHMMSSZ
                // LOCATION:___
                // SEQUENCE:___
                // STATUS:___
                // SUMMARY:___
                // TRANSP:___
                // X-GWSHOW-AS:___
                // END:VEVENT

                // Interpret line

                if (line.StartsWith("BEGIN:VEVENT"))
                {
                    // Create new event
                    ev = new Event();
                }
        
                else if (line.StartsWith("DTSTART:"))
                {
                    // Parse to integers and initialize start
                    int year  = int.Parse(line.Substring(8, 4));
                    int month = int.Parse(line.Substring(12, 2));
                    int day   = int.Parse(line.Substring(14, 2));
                    int hour  = int.Parse(line.Substring(17, 2));
                    int min   = int.Parse(line.Substring(19, 2));

                    ev.start = new DateTime(year, month, day, hour, min, 0);
                }
        
                else if (line.StartsWith("DTEND:"))
                {
                    // Parse values to integers
                    ev.endHour = int.Parse(line.Substring(15, 2));
                    ev.endMin  = int.Parse(line.Substring(17, 2));
                }
        
                /*
                else if (line.StartsWith("DTSTAMP:"))
                { }
        
                else if (line.StartsWith("UID:"))
                { }
        
                else if (line.StartsWith("CREATED:"))
                { }
        
                else if (line.StartsWith("LAST-MODIFIED:"))
                { }
                */
        
                else if (line.StartsWith("LOCATION:"))
                {
                    // Read location
                    ev.location = line.Substring(9);
                }
        
                /*
                else if (line.StartsWith("SEQUENCE:"))
                { }
        
                else if (line.StartsWith("STATUS:"))
                { }
                */
        
                else if (line.StartsWith("SUMMARY:"))
                {
                    ev.course      = ExtractFromSummary(line, SummaryToken.Course);
                    ev.teacher     = ExtractFromSummary(line, SummaryToken.Sign);
                    ev.description = ExtractFromSummary(line, SummaryToken.Description);

                    int number
                }
        
                /*
                else if (line.StartsWith("TRANSP:"))
                { }
        
                else if (line.StartsWith("X-GWSHOW-AS:"))
                { }
                */
        
                else if (line.StartsWith("END:VEVENT"))
                {
                    // Convert from UTC to local time

                    DateTime end = new DateTime(ev.start.Year, ev.start.Month, ev.start.Day, ev.endHour, ev.endMin, 0).ToLocalTime();
  
                    ev.endHour = end.Hour;
                    ev.endMin = end.Minute;
                    ev.week = calendar.GetWeekOfYear(ev.start, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

                    ev.start = ev.start.ToLocalTime();

                    // Write event to console

                    /*Console.WriteLine(
                        ev.week
                        + " | " + ev.start.ToString("ddd | yyyy-MM-dd | HH:mm")
                        +  "-"  + ev.endHour.ToString("00") + ":" + ev.endMin.ToString("00")
                        + " | " + ev.course
                        + " | " + ev.description
                        + " | " + ev.location);*/

                    // Add event to list
                    events.Add(ev);
                }

            } // End of loop

        } // End of method

    } // End of class

} // End of namespace

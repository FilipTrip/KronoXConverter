using System;

//////////////////////////////////////////
//
//  Event.cs
//  KronoX Converter by Filip Tripkovic
//  Last updated 2022-02-27
//
//////////////////////////////////////////

namespace KronoXConverter
{
    public struct Event
    {
        public DateTime start;
        public int week, endHour, endMin;
        public string course, location, description, teacher;
    };
}

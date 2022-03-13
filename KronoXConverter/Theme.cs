using System;
using System.Drawing;
using System.IO;
using System.Reflection;

//////////////////////////////////////////
//
//  Colors.cs
//  KronoX Converter by Filip Tripkovic
//  Last updated 2022-03-12
//
//////////////////////////////////////////

namespace KronoXConverter
{
    public static class Theme
    {
        public static Color fillHeader { get; private set; }
        public static Color fillBackground { get; private set; }
        public static Color fillScheduleLight { get; private set; }
        public static Color fillScheduleDark { get; private set; }

        public static Color fontColorHeader { get; private set; }
        public static Color fontColorPrimary { get; private set; }
        public static Color fontColorSecondary { get; private set; }
        public static Color fontColorCourse { get; private set; }
        public static Color fontColorSchedule { get; private set; }
        public static Color fontColorEnded { get; private set; }

        public static string typefaceHeading { get; private set; }
        public static string typefacePrimary { get; private set; }
        public static string updateButton { get; private set; }

        public static Color[] fillCourses { get; private set; }

        public static bool hasCourseColors { get; private set; }
        public static bool successful { get; private set; } // Whether or not the initialization was completed successfully
        public static string name { get; private set; } // The name of the theme

        public static void Initialize(string filePath)
        {
            successful = false;
            Theme.name = filePath.ToFileName(false);

            string type, name, value;
            Type theme = typeof(Theme);
            PropertyInfo property;

            foreach (string line in File.ReadAllLines(filePath))
            {
                if (line == "")
                    continue;

                try
                {
                    type = line.Split(' ')[0];
                    name = line.Split(' ')[1];
                    value = line.Substring(type.Length + name.Length + 2);
                    property = theme.GetProperty(name);

                    if (type == "color")
                    {
                        string[] components = value.Split(' ');

                        property.SetValue(null, Color.FromArgb(
                           int.Parse(components[0]),
                           int.Parse(components[1]),
                           int.Parse(components[2])));
                    }

                    else if (type == "string")
                    {
                        property.SetValue(null, value);
                    }

                    else if (type == "colors")
                    {
                        string[] values = value.Split(',');
                        string[] components;
                        Color[] colors = new Color[values.Length];
                        char[] sepparator = new char[] { ' ' };

                        for (int i = 0; i < values.Length; ++i)
                        {
                            components = values[i].Split(sepparator, StringSplitOptions.RemoveEmptyEntries);

                            colors[i] = Color.FromArgb(
                                int.Parse(components[0]),
                                int.Parse(components[1]),
                                int.Parse(components[2]));
                        }

                        property.SetValue(null, colors);
                    }
                }
                catch (Exception e)
                {
                    Program.WriteLine(Program.error, "Failed to read theme file: " + filePath.ToFileName(false));
                    Program.WriteLine(Program.secondary, "Cause: \"" + line + "\"");
                    Program.WriteLine(Program.secondary, "At: " + filePath);
                    Program.WriteLine(Program.secondary, "Message: " + e.Message);
                    return;
                }
            }

            hasCourseColors = fillCourses.Length > 0;
            successful = true;
        }
    }
}

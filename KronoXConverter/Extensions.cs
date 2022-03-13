using System;
using System.IO;
using System.Text.RegularExpressions;

//////////////////////////////////////////
//
//  Extensions.cs
//  KronoX Converter by Filip Tripkovic
//  Last updated 2022-03-02
//
//////////////////////////////////////////

namespace KronoXConverter
{
    public static class Extensions
    {
        // "oneOrMoreAnything (oneOrMoreDigits)"
        private static Regex regex = new Regex(".+ \\s \\( [0-9]+ \\)", RegexOptions.Singleline | RegexOptions.IgnorePatternWhitespace);

        /// <summary>
        /// Returns only the directory of this file path
        /// </summary>
        public static string ToDirectory(this string filePath)
        {
            return filePath.Remove(filePath.LastIndexOf('/'));
        }

        /// <summary>
        /// Returns only the file name of this file
        /// </summary>
        public static string ToFileName(this string filePath, bool includeExtension)
        {
            string fileName = filePath.Substring(filePath.LastIndexOf('/') + 1);

            if (includeExtension)
                return fileName;
            
            return fileName.Remove(fileName.LastIndexOf('.'));
        }

        /// <summary>
        /// Returns only the extension (including the dot) of this file
        /// </summary>
        public static string ToExtension(this string filePath)
        {
            return filePath.Substring(filePath.LastIndexOf('.'));
        }

        /// <summary>
        /// Returns a path to the first unique file name available
        /// </summary>
        public static string NextUniqeName(this string filePath)
        {
            string name = filePath.ToFileName(false);

            // If name matches "oneOrMoreAnything (oneOrMoreDigits)", remove " (oneOrMoreDigits)"
            if (regex.IsMatch(name))
                name = name.Remove(name.LastIndexOf(' '));

            // Add " (i)" to name until the file does not exist
            for (int i = 1; File.Exists(filePath); ++i)
                filePath = filePath.ToDirectory() + "/" + name + " (" + i + ")" + filePath.ToExtension();

            return filePath;
        }
    }
}

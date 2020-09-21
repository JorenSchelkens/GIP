using System;
using System.Collections.Generic;
using System.Linq;

namespace CalculationDomain.ErasmusHogeSchool
{
    public static class EHBFunctions
    {
        public static string FormatStringNonPercent(int toFormat)
        {
            if (toFormat == 0)
            {
                return "-";
            }
            return toFormat.ToString();
        }

        public static string FormatStringPercent(int toFormat)
        {
            if (toFormat == 0)
            {
                return "-";
            }
            return $"{toFormat.ToString()}%";
        }

        public static string FormatYearString(string currentYear, string nextYear)
        {
            return $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";
        }

        public static string FormatYearStringSpecial(string currentYear, string nextYear)
        {
            return $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}*";
        }

        public static string FormatYearDefaultString(string currentYear, string nextYear)
        {
            return $"{currentYear.Substring(2)}-{nextYear.Substring(2)}";
        }

        public static string GetFileOrderString()
        {
            List<string> result = new List<string>();

            DateTime current = DateTime.Now;

            if (DateTime.Now.Month <= 9)
            {
                current = current.AddYears(-1);
            }

            result.Add($"<b>Doorstroom</b> {(current.Year - 6).ToString().Substring(2)}-{(current.Year - 5).ToString().Substring(2)} <br>");
            result.Add($"<b>Doorstroom</b>  {(current.Year - 5).ToString().Substring(2)}-{(current.Year - 4).ToString().Substring(2)} <br>");
            result.Add($"<b>Doorstroom</b>  {(current.Year - 4).ToString().Substring(2)}-{(current.Year - 3).ToString().Substring(2)} <br>");
            result.Add($"<b>Doorstroom</b>  {(current.Year - 3).ToString().Substring(2)}-{(current.Year - 2).ToString().Substring(2)} <br>");
            result.Add($"<b>Doorstroom</b>  {(current.Year - 2).ToString().Substring(2)}-{(current.Year - 1).ToString().Substring(2)} <br>");

            result.Add($"<b>Instroom</b> {(current.Year - 5).ToString().Substring(2)}-{(current.Year - 4).ToString().Substring(2)} <br>");
            result.Add($"<b>Instroom</b> {(current.Year - 4).ToString().Substring(2)}-{(current.Year - 3).ToString().Substring(2)} <br>");
            result.Add($"<b>Instroom</b> {(current.Year - 3).ToString().Substring(2)}-{(current.Year - 2).ToString().Substring(2)} <br>");
            result.Add($"<b>Instroom</b> {(current.Year - 2).ToString().Substring(2)}-{(current.Year - 1).ToString().Substring(2)} <br>");
            result.Add($"<b>Instroom</b> {(current.Year - 1).ToString().Substring(2)}-{(current.Year).ToString().Substring(2)} <br>");

            result.Add($"<b>Uitstroom</b> {(current.Year - 6).ToString().Substring(2)}-{(current.Year - 5).ToString().Substring(2)} <br>");
            result.Add($"<b>Uitstroom</b> {(current.Year - 5).ToString().Substring(2)}-{(current.Year - 4).ToString().Substring(2)} <br>");
            result.Add($"<b>Uitstroom</b> {(current.Year - 4).ToString().Substring(2)}-{(current.Year - 3).ToString().Substring(2)} <br>");
            result.Add($"<b>Uitstroom</b> {(current.Year - 3).ToString().Substring(2)}-{(current.Year - 2).ToString().Substring(2)} <br>");
            result.Add($"<b>Uitstroom</b> {(current.Year - 2).ToString().Substring(2)}-{(current.Year - 1).ToString().Substring(2)} <br>");

            result.Add($"<b>EmptyPowerPoint</b>");

            return string.Concat(result);
        }

        public static DateTime GetCurrentAcademicYear()
        {
            DateTime current = DateTime.Now;

            if (DateTime.Now.Month <= 9)
            {
                current = current.AddYears(-1);
            }

            return current;
        }

        public static DateTime GetCurrentAcademicYearBasedOnIndex(int index)
        {
            DateTime current = DateTime.Now;

            if (DateTime.Now.Month <= 9)
            {
                current = current.AddYears(-1);
            }

            return current.AddYears(-(index - 1));
        }
    }
}
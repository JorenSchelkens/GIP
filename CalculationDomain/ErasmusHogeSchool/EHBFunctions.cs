using System;

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

        public static string FormatYearDefaultString(string currentYear, string nextYear)
        {
            return $"{currentYear.Substring(2)}-{nextYear.Substring(2)}";
        }

        public static DateTime GetCurrentAcademicYear()
        {
            DateTime current = DateTime.Now;

            if (DateTime.Now.Month < 9)
            {
                current = current.AddYears(-1);
            }

            return current;
        }
    }
}
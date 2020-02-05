using System;
using System.Collections.Generic;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class FilePathHandler
    {
        // @"d:\GitHub\GIP\";
        // @"C:\Users\joren.schelkens.BAZANDPOORT.000\Documents\GitHub\GIP\";

        private string DefaultAbsPath = @"d:\GitHub\GIP\";
        public List<string> DoorstroomPaths { get; set; } = new List<string>();
        public List<string> InstroomPaths { get; set; } = new List<string>();
        public List<string> UitstroomPaths { get; set; } = new List<string>();
        public string PowerPointPath { get; set; }
        public int MaxAantalPaths { get; set; } = 3;

        public FilePathHandler()
        {
            SetPaths();
        }

        private void SetPaths()
        {
            this.PowerPointPath = this.DefaultAbsPath + @"CalculationDomain\ErasmusHogeSchool\EmptyPowerPoint.pptx";

            DateTime currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = 0; i < this.MaxAantalPaths; i++)
            {
                this.InstroomPaths.Add(this.DefaultAbsPath +
                    @"CalculationDomain\ErasmusHogeSchool\Excels\Instroom " +
                    $"{EHBFunctions.FormatYearDefaultString(currentYear, nextYear)}.xlsx");

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }

            currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            currentYearTemp = currentYearTemp.AddYears(-1);
            currentYear = currentYearTemp.Year.ToString();
            nextYearTemp = currentYearTemp.AddYears(1);
            nextYear = nextYearTemp.Year.ToString();

            for (int i = 0; i < this.MaxAantalPaths; i++)
            {
                this.DoorstroomPaths.Add(this.DefaultAbsPath +
                    @"CalculationDomain\ErasmusHogeSchool\Excels\Doorstroom " +
                    $"{EHBFunctions.FormatYearDefaultString(currentYear, nextYear)}.xlsx");

                this.UitstroomPaths.Add(this.DefaultAbsPath +
                    @"CalculationDomain\ErasmusHogeSchool\Excels\Uitstroom " +
                    $"{EHBFunctions.FormatYearDefaultString(currentYear, nextYear)}.xlsx");

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }
        }
    }
}
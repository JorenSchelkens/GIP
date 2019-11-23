using DefaultDomain.ExcelReading;
using System.Collections.Generic;

namespace CalculationDomain.ErasmusHogeSchool.Uitstroom
{
    public class UitstroomBlad
    {
        public List<UitstroomRij> UitstroomRijen { get; set; } = new List<UitstroomRij>();

        public UitstroomBlad(string filePath, string opleiding)
        {
            List<Row> rows = Setup(filePath);
            FilterOpOpleiding(rows, opleiding);
        }

        public List<Row> Setup(string filePath)
        {
            return ExcelRead.ReadEHB(filePath);
        }

        public void FilterOpOpleiding(List<Row> rows, string opleiding)
        {
            foreach (Row row in rows)
            {
                if (row.columns[3] == opleiding)
                {
                    UitstroomRij uitstroomRij = new UitstroomRij();

                    uitstroomRij.SoOnderwijsvorm = row.columns[0];
                    uitstroomRij.Stamnummer = row.columns[1].Substring(0, 4);
                    uitstroomRij.DiplomaBehaald = (row.columns[2] == "Ja") ? true : false;

                    this.UitstroomRijen.Add(uitstroomRij);

                }
            }
        }
    }
}
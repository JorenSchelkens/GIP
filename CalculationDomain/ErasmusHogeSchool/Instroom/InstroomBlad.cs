using System.Collections.Generic;
using DefaultDomain.ExcelReading;

namespace CalculationDomain.ErasmusHogeSchool.Instroom
{
    public class InstroomBlad
    {
        public List<InstroomRij> InstroomRijen { get; set; } = new List<InstroomRij>();

        public InstroomBlad(string filePath, string opleiding)
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
                if(row.columns[4] == opleiding)
                {
                    InstroomRij instroomRij = new InstroomRij();

                    instroomRij.NieuweStudent = (row.columns[0] == "Ja") ? true : false;
                    instroomRij.GeneratieStudent = (row.columns[1] == "Ja") ? true : false;
                    instroomRij.SoOnderwijsvorm = row.columns[2];
                    //Traject

                    this.InstroomRijen.Add(instroomRij);

                }
            }
        }
    }
}
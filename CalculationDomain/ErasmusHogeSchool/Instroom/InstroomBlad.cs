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

                    if (row.columns[3].Contains("1/"))
                    {
                        int temp = row.columns[3].IndexOf("1/");
                        string sub = row.columns[3].Substring(temp + 2, 2);
                        instroomRij.Trajectschijfverdeling = int.Parse(sub);
                    }

                    this.InstroomRijen.Add(instroomRij);

                }
            }
        }

        public void FilterOpNieuweStudent()
        {
            List<InstroomRij> temp = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in this.InstroomRijen)
            {
                if (instroomRij.NieuweStudent)
                {
                    temp.Add(instroomRij);
                }
            }

            this.InstroomRijen = temp;
        }

        public List<InstroomRij> FilterOpVoltijds()
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in this.InstroomRijen)
            {
                if(instroomRij.Trajectschijfverdeling >= 54)
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen;
        }
    }
}
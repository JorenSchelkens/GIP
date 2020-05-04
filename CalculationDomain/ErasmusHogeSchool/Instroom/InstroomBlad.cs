using DefaultDomain.ExcelReading;
using System.Collections.Generic;

namespace CalculationDomain.ErasmusHogeSchool.Instroom
{
    public class InstroomBlad
    {
        public List<InstroomRij> InstroomRijen { get; set; } = new List<InstroomRij>();

        public InstroomBlad(string filePath, string opleiding)
        {
            List<Row> rows = this.Setup(filePath);
            this.FilterOpOpleiding(rows, opleiding);
        }

        public List<Row> Setup(string filePath)
        {
            return ExcelRead.ReadEHB(filePath);
        }

        public void FilterOpOpleiding(List<Row> rows, string opleiding)
        {
            foreach (Row row in rows)
            {
                if (row.columns[4] == opleiding)
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

        public List<InstroomRij> FilterOpNieuweStudent()
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in this.InstroomRijen)
            {
                if (instroomRij.NieuweStudent)
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen;
        }

        public List<InstroomRij> FilterOpVoltijds(List<InstroomRij> temp)
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in temp)
            {
                if (instroomRij.Trajectschijfverdeling >= 54)
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen;
        }

        public List<InstroomRij> FilterOpGeneratieStudent(List<InstroomRij> temp)
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in temp)
            {
                if (instroomRij.GeneratieStudent)
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen;
        }

        public int FilterOpASO(List<InstroomRij> temp)
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in temp)
            {
                if (instroomRij.SoOnderwijsvorm == "ASO" || instroomRij.SoOnderwijsvorm == "vASO")
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen.Count;
        }

        public int FilterOpTSO(List<InstroomRij> temp)
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in temp)
            {
                if (instroomRij.SoOnderwijsvorm == "TSO" || instroomRij.SoOnderwijsvorm == "vTSO")
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen.Count;
        }

        public int FilterOpBSO(List<InstroomRij> temp)
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in temp)
            {
                if (instroomRij.SoOnderwijsvorm == "BSO" || instroomRij.SoOnderwijsvorm == "vBSO")
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen.Count;
        }

        public int FilterOpKSO(List<InstroomRij> temp)
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in temp)
            {
                if (instroomRij.SoOnderwijsvorm == "KSO" || instroomRij.SoOnderwijsvorm == "vKSO")
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen.Count;
        }

        public int FilterOpAndereSO(List<InstroomRij> temp)
        {
            List<InstroomRij> instroomRijen = new List<InstroomRij>();

            foreach (InstroomRij instroomRij in temp)
            {
                if (instroomRij.SoOnderwijsvorm != "ASO" &&
                    instroomRij.SoOnderwijsvorm != "vASO" &&
                    instroomRij.SoOnderwijsvorm != "TSO" &&
                    instroomRij.SoOnderwijsvorm != "vTSO" &&
                    instroomRij.SoOnderwijsvorm != "BSO" &&
                    instroomRij.SoOnderwijsvorm != "vBSO" &&
                    instroomRij.SoOnderwijsvorm != "KSO" &&
                    instroomRij.SoOnderwijsvorm != "vKSO")
                {
                    instroomRijen.Add(instroomRij);
                }
            }

            return instroomRijen.Count;
        }
    }
}
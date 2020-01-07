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

        public List<UitstroomRij> FilterHeeftDiploma(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.DiplomaBehaald)
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen;
        }

        public int FilterOpASO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "ASO" || uitstroomRij.SoOnderwijsvorm == "vASO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen.Count;
        }

        public int FilterOpTSO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "TSO" || uitstroomRij.SoOnderwijsvorm == "vTSO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen.Count;
        }

        public int FilterOpBSO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "BSO" || uitstroomRij.SoOnderwijsvorm == "vBSO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen.Count;
        }

        public int FilterOpKSO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "KSO" || uitstroomRij.SoOnderwijsvorm == "vKSO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen.Count;
        }

        public int FilterOpAndereSO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm != "ASO" &&
                    uitstroomRij.SoOnderwijsvorm != "vASO" &&
                    uitstroomRij.SoOnderwijsvorm != "TSO" &&
                    uitstroomRij.SoOnderwijsvorm != "vTSO" &&
                    uitstroomRij.SoOnderwijsvorm != "BSO" &&
                    uitstroomRij.SoOnderwijsvorm != "vBSO" &&
                    uitstroomRij.SoOnderwijsvorm != "KSO" &&
                    uitstroomRij.SoOnderwijsvorm != "vKSO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen.Count;
        }
    }
}
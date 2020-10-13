using DefaultDomain.ExcelReading;
using System;
using System.Collections.Generic;
using System.IO;

namespace CalculationDomain.ErasmusHogeSchool.Uitstroom
{
    public class UitstroomBlad
    {
        public List<UitstroomRij> UitstroomRijen { get; set; } = new List<UitstroomRij>();

        public UitstroomBlad(MemoryStream stream, string opleiding)
        {
            List<Row> rows = this.Setup(stream);
            this.FilterOpOpleiding(rows, opleiding);
        }

        public List<Row> Setup(Stream stream)
        {
            return ExcelRead.ReadEHB(stream);
        }

        public void FilterOpOpleiding(List<Row> rows, string opleiding)
        {
            foreach (Row row in rows)
            {
                if (!string.IsNullOrEmpty(row.columns[3]))
                {
                    if (row.columns[3].ToLower() == opleiding.ToLower())
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

        public int FilterOpMinderDan3(List<UitstroomRij> temp, DateTime currentYear)
        {
            int total = 0;
            DateTime date;
            DateTime difference;

            foreach (UitstroomRij uitstroomRij in temp)
            {
                date = new DateTime(int.Parse(uitstroomRij.Stamnummer), 1, 1);
                difference = currentYear.AddYears(-(date.Year));

                if (difference.Year < 3)
                {
                    total++;
                }
            }

            return total;
        }

        public int FilterOp3(List<UitstroomRij> temp, DateTime currentYear)
        {
            int total = 0;
            DateTime date;
            DateTime difference;

            foreach (UitstroomRij uitstroomRij in temp)
            {
                date = new DateTime(int.Parse(uitstroomRij.Stamnummer), 1, 1);
                difference = currentYear.AddYears(-(date.Year));

                if (difference.Year == 3)
                {
                    total++;
                }
            }

            return total;
        }

        public int FilterOpMeerDan3(List<UitstroomRij> temp, DateTime currentYear)
        {
            int total = 0;
            DateTime date;
            DateTime difference;

            foreach (UitstroomRij uitstroomRij in temp)
            {
                date = new DateTime(int.Parse(uitstroomRij.Stamnummer), 1, 1);
                difference = currentYear.AddYears(-(date.Year));

                if (difference.Year > 3)
                {
                    total++;
                }
            }

            return total;
        }

        public int FilterOp4(List<UitstroomRij> temp, DateTime currentYear)
        {
            int total = 0;
            DateTime date;
            DateTime difference;

            foreach (UitstroomRij uitstroomRij in temp)
            {
                date = new DateTime(int.Parse(uitstroomRij.Stamnummer), 1, 1);
                difference = currentYear.AddYears(-(date.Year));

                if (difference.Year == 4)
                {
                    total++;
                }
            }

            return total;
        }

        public int FilterOpMeerDan4(List<UitstroomRij> temp, DateTime currentYear)
        {
            int total = 0;
            DateTime date;
            DateTime difference;

            foreach (UitstroomRij uitstroomRij in temp)
            {
                date = new DateTime(int.Parse(uitstroomRij.Stamnummer), 1, 1);
                difference = currentYear.AddYears(-(date.Year));

                if (difference.Year > 4)
                {
                    total++;
                }
            }

            return total;
        }

        public List<UitstroomRij> FilterOpASO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "ASO" || uitstroomRij.SoOnderwijsvorm == "vASO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen;
        }

        public List<UitstroomRij> FilterOpTSO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "TSO" || uitstroomRij.SoOnderwijsvorm == "vTSO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen;
        }

        public List<UitstroomRij> FilterOpBSO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "BSO" || uitstroomRij.SoOnderwijsvorm == "vBSO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen;
        }

        public List<UitstroomRij> FilterOpKSO(List<UitstroomRij> temp)
        {
            List<UitstroomRij> uitstroomRijen = new List<UitstroomRij>();

            foreach (UitstroomRij uitstroomRij in temp)
            {
                if (uitstroomRij.SoOnderwijsvorm == "KSO" || uitstroomRij.SoOnderwijsvorm == "vKSO")
                {
                    uitstroomRijen.Add(uitstroomRij);
                }
            }

            return uitstroomRijen;
        }

        public List<UitstroomRij> FilterOpAndereSO(List<UitstroomRij> temp)
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

            return uitstroomRijen;
        }
    }
}
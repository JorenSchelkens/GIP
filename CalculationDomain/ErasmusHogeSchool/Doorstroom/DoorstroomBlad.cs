using DefaultDomain.ExcelReading;
using System.Collections.Generic;

namespace CalculationDomain.ErasmusHogeSchool.Doorstroom
{
    public class DoorstroomBlad
    {
        public List<DoorstroomRij> DoorstroomRijen { get; set; } = new List<DoorstroomRij>();

        public DoorstroomBlad(string filePath, string opleiding)
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
                if (row.columns[10] == opleiding)
                {
                    DoorstroomRij doorstroomRij = new DoorstroomRij();

                    doorstroomRij.StudiepuntenTeVolgen = int.Parse(row.columns[0]);
                    doorstroomRij.StudiepuntenCredits = int.Parse(row.columns[1]);
                    doorstroomRij.VolgtOlodInSchijf1 = (row.columns[2] == "Ja") ? true : false;
                    doorstroomRij.NieuweStudentInInstelling = (row.columns[3] == "Ja") ? true : false;

                    if (row.columns[4].Contains("1/"))
                    {
                        int temp = row.columns[4].IndexOf("1/");
                        string sub = row.columns[4].Substring(temp + 2, 2);
                        doorstroomRij.Trajectschijfverdeling = int.Parse(sub);
                    }

                    doorstroomRij.SoOnderwijsvorm = row.columns[5];
                    doorstroomRij.Stamnummer = row.columns[6].Substring(0, 4);
                    doorstroomRij.KanDiplomaBehalen = row.columns[7];
                    doorstroomRij.HeeftDiplomaBehaalt = (row.columns[8] == "Ja") ? true : false;
                    doorstroomRij.Generatie = (row.columns[9] == "Ja") ? true : false;

                    this.DoorstroomRijen.Add(doorstroomRij);

                }
            }
        }

        public List<DoorstroomRij> FilterOpVoltijds(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.Trajectschijfverdeling >= 54)
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpOLOD(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.VolgtOlodInSchijf1)
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpNieuweStudent(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.NieuweStudentInInstelling)
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOp60Stp(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.StudiepuntenCredits == 60)
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpTussen60En45Stp(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.StudiepuntenCredits >= 45 && doorstroomRij.StudiepuntenCredits < 60)
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpMinderDan45Stp(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.StudiepuntenCredits < 45)
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpASO(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.SoOnderwijsvorm == "ASO" || doorstroomRij.SoOnderwijsvorm == "vASO")
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpTSO(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.SoOnderwijsvorm == "TSO" || doorstroomRij.SoOnderwijsvorm == "vTSO")
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpBSO(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.SoOnderwijsvorm == "BSO" || doorstroomRij.SoOnderwijsvorm == "vBSO")
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpKSO(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.SoOnderwijsvorm == "KSO" || doorstroomRij.SoOnderwijsvorm == "vKSO")
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }

        public List<DoorstroomRij> FilterOpAndereSO(List<DoorstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomRijen = new List<DoorstroomRij>();

            foreach (DoorstroomRij doorstroomRij in temp)
            {
                if (doorstroomRij.SoOnderwijsvorm != "ASO" &&
                    doorstroomRij.SoOnderwijsvorm != "vASO" &&
                    doorstroomRij.SoOnderwijsvorm != "TSO" &&
                    doorstroomRij.SoOnderwijsvorm != "vTSO" &&
                    doorstroomRij.SoOnderwijsvorm != "BSO" &&
                    doorstroomRij.SoOnderwijsvorm != "vBSO" &&
                    doorstroomRij.SoOnderwijsvorm != "KSO" &&
                    doorstroomRij.SoOnderwijsvorm != "vKSO")
                {
                    doorstroomRijen.Add(doorstroomRij);
                }
            }

            return doorstroomRijen;
        }
    }
}
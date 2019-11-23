using DefaultDomain.ExcelReading;
using System.Collections.Generic;

namespace CalculationDomain.ErasmusHogeSchool.Doorstroom
{
    public class DoorstroomBlad
    {
        public List<DoorstroomRij> DoorstroomRijen { get; set; } = new List<DoorstroomRij>();

        public DoorstroomBlad(string filePath, string opleiding)
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
                if (row.columns[10] == opleiding)
                {
                    DoorstroomRij doorstroomRij = new DoorstroomRij();

                    doorstroomRij.StudiepuntenTeVolgen = int.Parse(row.columns[0]);
                    doorstroomRij.StudiepuntenCredits = int.Parse(row.columns[1]);
                    doorstroomRij.VolgtOlodInSchijf1 = (row.columns[2] == "Ja") ? true : false;
                    doorstroomRij.NieuweStudentInInstelling = (row.columns[3] == "Ja") ? true : false;
                    //Traject
                    doorstroomRij.SoOnderwijsvorm = row.columns[5];
                    doorstroomRij.Stamnummer = row.columns[6].Substring(0, 4);
                    doorstroomRij.KanDiplomaBehalen = row.columns[7];
                    doorstroomRij.HeeftDiplomaBehaalt = (row.columns[8] == "Ja") ? true : false;
                    doorstroomRij.Generatie = (row.columns[9] == "Ja") ? true : false;

                    this.DoorstroomRijen.Add(doorstroomRij);

                }
            }
        }
    }
}
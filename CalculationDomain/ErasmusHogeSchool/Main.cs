using CalculationDomain.ErasmusHogeSchool.Instroom;
using CalculationDomain.ErasmusHogeSchool.Doorstroom;
using CalculationDomain.ErasmusHogeSchool.Uitstroom;
using System.Collections.Generic;
using System;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class Main
    {
        public InstroomBlad InstroomBlad { get; set; }
        public DoorstroomBlad DoorstroomBlad { get; set; }
        public UitstroomBlad UitstroomBlad { get; set; }
        public PowerPointClass PowerPoint { get; set; }
        private string Filter { get; set; }

        public Main(string opleiding)
        {
            this.Filter = opleiding;
            this.PowerPoint = new PowerPointClass(opleiding);
        }

        public void Load()
        {
            //Haal bestanden

            // d:\GitHub\GIP\Documentation\
            // C:\Users\joren.schelkens.BAZANDPOORT.000\Documents\GitHub\GIP\Documentation\

            this.InstroomBlad = new InstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Instroom dd.08.10.2019 - 19-20.xlsx", this.Filter);
            this.DoorstroomBlad = new DoorstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Doorstroom dd.08.10.2019 -18-19.xlsx", this.Filter);
            this.UitstroomBlad = new UitstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Uitstroom dd.08.10.2019 - 18-19.xlsx", this.Filter);
        }

        public void GenerateInstroomData()
        {
            this.InstroomBlad.FilterOpNieuweStudent();
            List<InstroomRij> instroomVoltijdsTemp = this.InstroomBlad.FilterOpVoltijds();

            int deeltijds = this.InstroomBlad.InstroomRijen.Count - instroomVoltijdsTemp.Count;
            int totaal = this.InstroomBlad.InstroomRijen.Count;

            List<InstroomRij> instroomGeneratieStudentTemp = this.InstroomBlad.FilterOpGeneratieStudent();
            int generatiestudent = instroomGeneratieStudentTemp.Count;

            int nietGeneratiestudent = this.InstroomBlad.InstroomRijen.Count - generatiestudent;

            double aandelInTotaal = ((double) generatiestudent / totaal) * 100;
            double aandeelInVoltijds = ((double) generatiestudent / instroomVoltijdsTemp.Count) * 100;

            this.PowerPoint.AddInstroomSlide(
                instroomVoltijdsTemp.Count, 
                deeltijds, 
                totaal, 
                generatiestudent, 
                nietGeneratiestudent,
                (int)Math.Round(aandelInTotaal),
                (int)Math.Round(aandeelInVoltijds));
        }

        public void SavePowerPoint()
        {
            this.PowerPoint.Save();
        }
    }
}
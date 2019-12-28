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

        public void GenerateInstroomData1()
        {
            List<InstroomRij> instroomNieuweStudenten = this.InstroomBlad.FilterOpNieuweStudent();
            List<InstroomRij> instroomVoltijdsTemp = this.InstroomBlad.FilterOpVoltijds(instroomNieuweStudenten);

            int deeltijds = instroomNieuweStudenten.Count - instroomVoltijdsTemp.Count;
            int totaal = instroomNieuweStudenten.Count;

            List<InstroomRij> instroomGeneratieStudentTemp = this.InstroomBlad.FilterOpGeneratieStudent(instroomNieuweStudenten);
            int generatiestudent = instroomGeneratieStudentTemp.Count;

            int nietGeneratiestudent = instroomNieuweStudenten.Count - generatiestudent;

            double aandelInTotaal = ((double) generatiestudent / totaal) * 100;
            double aandeelInVoltijds = ((double) generatiestudent / instroomVoltijdsTemp.Count) * 100;

            this.PowerPoint.AddInstroomSlide1(
                instroomVoltijdsTemp.Count, 
                deeltijds, 
                totaal, 
                generatiestudent, 
                nietGeneratiestudent,
                (int)Math.Round(aandelInTotaal),
                (int)Math.Round(aandeelInVoltijds));
        }

        public void GenerateInstroomData2()
        {
            List<InstroomRij> instroomNieuweStudenten = this.InstroomBlad.FilterOpNieuweStudent();
            List<InstroomRij> instroomVoltijdsTemp = this.InstroomBlad.FilterOpVoltijds(instroomNieuweStudenten);

            this.PowerPoint.AddInstroomSlide2(
                this.InstroomBlad.FilterOpASO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpTSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpBSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpKSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpAndereSO(instroomVoltijdsTemp),
                instroomVoltijdsTemp.Count);

            double aandelASO = ((double)this.InstroomBlad.FilterOpASO(instroomVoltijdsTemp) / instroomVoltijdsTemp.Count) * 100;
            double aandelTSO = ((double)this.InstroomBlad.FilterOpTSO(instroomVoltijdsTemp) / instroomVoltijdsTemp.Count) * 100;
            double aandelBSO = ((double)this.InstroomBlad.FilterOpBSO(instroomVoltijdsTemp) / instroomVoltijdsTemp.Count) * 100;
            double aandelKSO = ((double)this.InstroomBlad.FilterOpKSO(instroomVoltijdsTemp) / instroomVoltijdsTemp.Count) * 100;
            double aandelAndereSO = ((double)this.InstroomBlad.FilterOpAndereSO(instroomVoltijdsTemp) / instroomVoltijdsTemp.Count) * 100;

            this.PowerPoint.AddInstroomSlide3(
                (int)Math.Round(aandelASO),
                (int)Math.Round(aandelTSO),
                (int)Math.Round(aandelBSO),
                (int)Math.Round(aandelKSO),
                (int)Math.Round(aandelAndereSO));
        }

        public void SavePowerPoint()
        {
            this.PowerPoint.Save();
        }
    }
}
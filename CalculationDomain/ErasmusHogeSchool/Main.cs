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
            // C:\Users\Joren\Documents\GitHub\GIP\Documentation\

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

            this.PowerPoint.ChangeInstroomSlide1(
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

            this.PowerPoint.ChangeInstroomSlide2(
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

            this.PowerPoint.ChangeInstroomSlide3(
                (int)Math.Round(aandelASO),
                (int)Math.Round(aandelTSO),
                (int)Math.Round(aandelBSO),
                (int)Math.Round(aandelKSO),
                (int)Math.Round(aandelAndereSO));
        }

        public void GenerateInstroomData3()
        {
            List<InstroomRij> instroomGeneratieTemp = this.InstroomBlad.FilterOpGeneratieStudent(this.InstroomBlad.InstroomRijen);
            List<InstroomRij> instroomVoltijdsTemp = this.InstroomBlad.FilterOpVoltijds(instroomGeneratieTemp);

            this.PowerPoint.ChangeInstroomSlide4(
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

            this.PowerPoint.ChangeInstroomSlide5(
                (int)Math.Round(aandelASO),
                (int)Math.Round(aandelTSO),
                (int)Math.Round(aandelBSO),
                (int)Math.Round(aandelKSO),
                (int)Math.Round(aandelAndereSO));
        }

        public void GenerateDoorstroomData1()
        {

        }

        public void GenerateUitstroomData1()
        {
            List<UitstroomRij> uitstroomDiplomaTemp = this.UitstroomBlad.FilterHeeftDiploma(this.UitstroomBlad.UitstroomRijen);

            double aandelASO = ((double)this.UitstroomBlad.FilterOpASO(uitstroomDiplomaTemp) / uitstroomDiplomaTemp.Count) * 100;
            double aandelTSO = ((double)this.UitstroomBlad.FilterOpTSO(uitstroomDiplomaTemp) / uitstroomDiplomaTemp.Count) * 100;
            double aandelBSO = ((double)this.UitstroomBlad.FilterOpBSO(uitstroomDiplomaTemp) / uitstroomDiplomaTemp.Count) * 100;
            double aandelKSO = ((double)this.UitstroomBlad.FilterOpKSO(uitstroomDiplomaTemp) / uitstroomDiplomaTemp.Count) * 100;
            double aandelAndereSO = ((double)this.UitstroomBlad.FilterOpAndereSO(uitstroomDiplomaTemp) / uitstroomDiplomaTemp.Count) * 100;

            this.PowerPoint.ChangeUitstroomSlide1(
                this.UitstroomBlad.FilterOpASO(uitstroomDiplomaTemp),
                this.UitstroomBlad.FilterOpTSO(uitstroomDiplomaTemp),
                this.UitstroomBlad.FilterOpBSO(uitstroomDiplomaTemp),
                this.UitstroomBlad.FilterOpKSO(uitstroomDiplomaTemp),
                this.UitstroomBlad.FilterOpAndereSO(uitstroomDiplomaTemp),
                uitstroomDiplomaTemp.Count,
                (int)Math.Round(aandelASO),
                (int)Math.Round(aandelTSO),
                (int)Math.Round(aandelBSO),
                (int)Math.Round(aandelKSO),
                (int)Math.Round(aandelAndereSO));
        }

        public void GenerateAll()
        {
            Load();

            GenerateInstroomData1();
            GenerateInstroomData2();
            GenerateInstroomData3();

            //GenerateDoorstroomData1();

            GenerateUitstroomData1();
        }

        public void SavePowerPoint()
        {
            this.PowerPoint.Save();
        }
    }
}
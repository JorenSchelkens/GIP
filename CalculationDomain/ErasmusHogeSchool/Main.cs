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
        private FilePathHandler FilePathHandler { get; set; } = new FilePathHandler();
        private string Filter { get; set; }

        public Main(string opleiding)
        {
            this.Filter = opleiding;
            this.PowerPoint = new PowerPointClass(opleiding);
        }

        private void Load(int index)
        {
            this.InstroomBlad = new InstroomBlad(this.FilePathHandler.InstroomPaths[index], this.Filter);
            this.DoorstroomBlad = new DoorstroomBlad(this.FilePathHandler.DoorstroomPaths[index], this.Filter);
            this.UitstroomBlad = new UitstroomBlad(this.FilePathHandler.UitstroomPaths[index], this.Filter);
        }

        private void GenerateInstroomData1(int index)
        {
            List<InstroomRij> instroomNieuweStudenten = this.InstroomBlad.FilterOpNieuweStudent();
            List<InstroomRij> instroomVoltijdsTemp = this.InstroomBlad.FilterOpVoltijds(instroomNieuweStudenten);

            int deeltijds = instroomNieuweStudenten.Count - instroomVoltijdsTemp.Count;
            int totaal = instroomNieuweStudenten.Count;

            List<InstroomRij> instroomGeneratieStudentTemp = this.InstroomBlad.FilterOpGeneratieStudent(instroomNieuweStudenten);
            int generatiestudent = instroomGeneratieStudentTemp.Count;

            int nietGeneratiestudent = instroomNieuweStudenten.Count - generatiestudent;

            List<InstroomRij> instroomVoltijdsEnGeneratieTemp = this.InstroomBlad.FilterOpVoltijds(instroomGeneratieStudentTemp);
            int instroomVoltijdsEnGeneratie = instroomVoltijdsEnGeneratieTemp.Count;

            double aandelInTotaal = ((double) generatiestudent / totaal) * 100;
            double aandeelInVoltijds = ((double)instroomVoltijdsEnGeneratie / instroomVoltijdsTemp.Count) * 100;

            this.PowerPoint.ChangeInstroomSlide1(
                instroomVoltijdsTemp.Count, 
                deeltijds, 
                totaal, 
                generatiestudent, 
                nietGeneratiestudent,
                (int)Math.Round(aandelInTotaal),
                (int)Math.Round(aandeelInVoltijds),
                index);
        }

        private void GenerateInstroomData2(int index)
        {
            List<InstroomRij> instroomNieuweStudenten = this.InstroomBlad.FilterOpNieuweStudent();
            List<InstroomRij> instroomVoltijdsTemp = this.InstroomBlad.FilterOpVoltijds(instroomNieuweStudenten);

            this.PowerPoint.ChangeInstroomSlide2(
                this.InstroomBlad.FilterOpASO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpTSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpBSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpKSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpAndereSO(instroomVoltijdsTemp),
                instroomVoltijdsTemp.Count,
                index);

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
                (int)Math.Round(aandelAndereSO),
                index);
        }

        private void GenerateInstroomData3(int index)
        {
            List<InstroomRij> instroomGeneratieTemp = this.InstroomBlad.FilterOpGeneratieStudent(this.InstroomBlad.InstroomRijen);
            List<InstroomRij> instroomVoltijdsTemp = this.InstroomBlad.FilterOpVoltijds(instroomGeneratieTemp);

            this.PowerPoint.ChangeInstroomSlide4(
                this.InstroomBlad.FilterOpASO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpTSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpBSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpKSO(instroomVoltijdsTemp),
                this.InstroomBlad.FilterOpAndereSO(instroomVoltijdsTemp),
                instroomVoltijdsTemp.Count,
                index);

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
                (int)Math.Round(aandelAndereSO),
                index);
        }

        private void GenerateDoorstroomData1(int index)
        {

        }

        private void GenerateUitstroomData1(int index)
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
                (int)Math.Round(aandelAndereSO),
                index);
        }

        public void GenerateAll()
        {
            for (int i = 0; i < FilePathHandler.MaxAantalPaths; i++)
            {
                Load(i);

                GenerateInstroomData1(i + 1);
                GenerateInstroomData2(i + 1);
                GenerateInstroomData3(i + 1);

                GenerateDoorstroomData1(i + 1);

                GenerateUitstroomData1(i + 1);
            }
        }

        public void SavePowerPoint()
        {
            this.PowerPoint.Save();
        }
    }
}
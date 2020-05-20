using CalculationDomain.ErasmusHogeSchool.Doorstroom;
using CalculationDomain.ErasmusHogeSchool.Instroom;
using CalculationDomain.ErasmusHogeSchool.Uitstroom;
using System;
using System.Collections.Generic;

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

        private List<InstroomRij> GenerateInstroomData1(int index)
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

            double aandelInTotaal = ((double)generatiestudent / totaal) * 100;
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

            return instroomVoltijdsTemp;
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

        private void GenerateDoorstroomData1(int index, List<InstroomRij> temp)
        {
            List<DoorstroomRij> doorstroomOLODTemp = this.DoorstroomBlad.FilterOpOLOD(this.DoorstroomBlad.DoorstroomRijen);
            List<DoorstroomRij> doorstroomOLODEnNieweStudentTemp = this.DoorstroomBlad.FilterOpNieuweStudent(doorstroomOLODTemp);
            List<DoorstroomRij> doorstroomOLODEnNieweStudentEnVoltijdsTemp = this.DoorstroomBlad.FilterOpVoltijds(doorstroomOLODEnNieweStudentTemp);

            List<DoorstroomRij> doorstroom60 = this.DoorstroomBlad.FilterOp60Stp(doorstroomOLODEnNieweStudentEnVoltijdsTemp);
            List<DoorstroomRij> doorstroomMeer45 = this.DoorstroomBlad.FilterOpTussen60En45Stp(doorstroomOLODEnNieweStudentEnVoltijdsTemp);
            List<DoorstroomRij> doorstroomMinder45 = this.DoorstroomBlad.FilterOpMinderDan45Stp(doorstroomOLODEnNieweStudentEnVoltijdsTemp);

            double zestigStp1 = ((double)doorstroom60.Count / temp.Count) * 100;
            double tussenZestigStpEnVijfenveertig1 = ((double)doorstroomMeer45.Count / temp.Count) * 100;
            double onderVijfenveertig1 = ((double)doorstroomMinder45.Count / temp.Count) * 100;
            double dropOut = 0;



            double zestigStp2 = ((double)doorstroom60.Count / doorstroomOLODEnNieweStudentEnVoltijdsTemp.Count) * 100;
            double tussenZestigStpEnVijfenveertig2 = ((double)doorstroomMeer45.Count / doorstroomOLODEnNieweStudentEnVoltijdsTemp.Count) * 100;
            double onderVijfenveertig2 = ((double)doorstroomMinder45.Count / doorstroomOLODEnNieweStudentEnVoltijdsTemp.Count) * 100;

            this.PowerPoint.ChangeDoorstroomSlide1(
                (int)Math.Round(zestigStp1),
                (int)Math.Round(tussenZestigStpEnVijfenveertig1),
                (int)Math.Round(onderVijfenveertig1),
                (int)Math.Round(dropOut),
                (int)Math.Round(zestigStp2),
                (int)Math.Round(tussenZestigStpEnVijfenveertig2),
                (int)Math.Round(onderVijfenveertig2),
                index);
        }

        private void GenerateDoorstroomData2()
        {
            List<DoorstroomRij> doorstroomNieweStudentTemp = this.DoorstroomBlad.FilterOpNieuweStudent(this.DoorstroomBlad.DoorstroomRijen);
            List<DoorstroomRij> doorstroomNieweStudentEnVoltijdsTemp = this.DoorstroomBlad.FilterOpVoltijds(doorstroomNieweStudentTemp);

            List<DoorstroomRij> doorstroomAso = this.DoorstroomBlad.FilterOpASO(doorstroomNieweStudentEnVoltijdsTemp);
            List<DoorstroomRij> doorstroomTso = this.DoorstroomBlad.FilterOpTSO(doorstroomNieweStudentEnVoltijdsTemp);
            List<DoorstroomRij> doorstroomKso = this.DoorstroomBlad.FilterOpKSO(doorstroomNieweStudentEnVoltijdsTemp);
            List<DoorstroomRij> doorstroomBso = this.DoorstroomBlad.FilterOpBSO(doorstroomNieweStudentEnVoltijdsTemp);
            List<DoorstroomRij> doorstroomAndereSO = this.DoorstroomBlad.FilterOpAndereSO(doorstroomNieweStudentEnVoltijdsTemp);

            this.PowerPoint.ChangeDoorstroomSlide2(
                (int)Math.Round((doorstroomAso.Count != 0) ? (double)this.DoorstroomBlad.FilterOp60Stp(doorstroomAso).Count / doorstroomAso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomAso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpTussen60En45Stp(doorstroomAso).Count / doorstroomAso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomAso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpMinderDan45Stp(doorstroomAso).Count / doorstroomAso.Count * 100.00 : 0.00),
                doorstroomAso.Count,
                (int)Math.Round((doorstroomTso.Count != 0) ? (double)this.DoorstroomBlad.FilterOp60Stp(doorstroomTso).Count / doorstroomTso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomTso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpTussen60En45Stp(doorstroomTso).Count / doorstroomTso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomTso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpMinderDan45Stp(doorstroomTso).Count / doorstroomTso.Count * 100.00 : 0.00),
                doorstroomTso.Count,
                (int)Math.Round((doorstroomKso.Count != 0) ? (double)this.DoorstroomBlad.FilterOp60Stp(doorstroomKso).Count / doorstroomKso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomKso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpTussen60En45Stp(doorstroomKso).Count / doorstroomKso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomKso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpMinderDan45Stp(doorstroomKso).Count / doorstroomKso.Count * 100.00 : 0.00),
                doorstroomKso.Count,
                (int)Math.Round((doorstroomKso.Count != 0) ? (double)this.DoorstroomBlad.FilterOp60Stp(doorstroomBso).Count / doorstroomBso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomKso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpTussen60En45Stp(doorstroomBso).Count / doorstroomBso.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomKso.Count != 0) ? (double)this.DoorstroomBlad.FilterOpMinderDan45Stp(doorstroomBso).Count / doorstroomBso.Count * 100.00 : 0.00),
                doorstroomKso.Count,
                (int)Math.Round((doorstroomAndereSO.Count != 0) ? (double)this.DoorstroomBlad.FilterOp60Stp(doorstroomAndereSO).Count / doorstroomAndereSO.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomAndereSO.Count != 0) ? (double)this.DoorstroomBlad.FilterOpTussen60En45Stp(doorstroomAndereSO).Count / doorstroomAndereSO.Count * 100.00 : 0.00),
                (int)Math.Round((doorstroomAndereSO.Count != 0) ? (double)this.DoorstroomBlad.FilterOpMinderDan45Stp(doorstroomAndereSO).Count / doorstroomAndereSO.Count * 100.00 : 0.00),
                doorstroomAndereSO.Count);
        }

        private void GenerateUitstroomData1(int index)
        {
            List<UitstroomRij> uitstroomDiplomaTemp = this.UitstroomBlad.FilterHeeftDiploma(this.UitstroomBlad.UitstroomRijen);

            double aandelASO = ((double)this.UitstroomBlad.FilterOpASO(uitstroomDiplomaTemp).Count / uitstroomDiplomaTemp.Count) * 100;
            double aandelTSO = ((double)this.UitstroomBlad.FilterOpTSO(uitstroomDiplomaTemp).Count / uitstroomDiplomaTemp.Count) * 100;
            double aandelBSO = ((double)this.UitstroomBlad.FilterOpBSO(uitstroomDiplomaTemp).Count / uitstroomDiplomaTemp.Count) * 100;
            double aandelKSO = ((double)this.UitstroomBlad.FilterOpKSO(uitstroomDiplomaTemp).Count / uitstroomDiplomaTemp.Count) * 100;
            double aandelAndereSO = ((double)this.UitstroomBlad.FilterOpAndereSO(uitstroomDiplomaTemp).Count / uitstroomDiplomaTemp.Count) * 100;

            this.PowerPoint.ChangeUitstroomSlide1(
                this.UitstroomBlad.FilterOpASO(uitstroomDiplomaTemp).Count,
                this.UitstroomBlad.FilterOpTSO(uitstroomDiplomaTemp).Count,
                this.UitstroomBlad.FilterOpBSO(uitstroomDiplomaTemp).Count,
                this.UitstroomBlad.FilterOpKSO(uitstroomDiplomaTemp).Count,
                this.UitstroomBlad.FilterOpAndereSO(uitstroomDiplomaTemp).Count,
                uitstroomDiplomaTemp.Count,
                (int)Math.Round(aandelASO),
                (int)Math.Round(aandelTSO),
                (int)Math.Round(aandelBSO),
                (int)Math.Round(aandelKSO),
                (int)Math.Round(aandelAndereSO),
                index);
        }

        private void GenerateStudieduur1(int index)
        {
            List<UitstroomRij> uitstroomDiploma = this.UitstroomBlad.FilterHeeftDiploma(this.UitstroomBlad.UitstroomRijen);

            int minderDanDrie = this.UitstroomBlad.FilterOpMinderDan3(uitstroomDiploma, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index));
            int drie = this.UitstroomBlad.FilterOp3(uitstroomDiploma, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index));
            int vier = this.UitstroomBlad.FilterOp4(uitstroomDiploma, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index));
            int meerDanVier = this.UitstroomBlad.FilterOpMeerDan4(uitstroomDiploma, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index));

            this.PowerPoint.ChangeStudieduurSlide1(
                (int)Math.Round((uitstroomDiploma.Count != 0) ? (double)minderDanDrie / uitstroomDiploma.Count * 100.00 : 0.00),
                (int)Math.Round((uitstroomDiploma.Count != 0) ? (double)drie / uitstroomDiploma.Count * 100.00 : 0.00),
                (int)Math.Round((uitstroomDiploma.Count != 0) ? (double)vier / uitstroomDiploma.Count * 100.00 : 0.00),
                (int)Math.Round((uitstroomDiploma.Count != 0) ? (double)meerDanVier / uitstroomDiploma.Count * 100.00 : 0.00),
                index);
        }

        private void GenerateStudieduur2(int index)
        {
            List<UitstroomRij> uitstroomDiploma = this.UitstroomBlad.FilterHeeftDiploma(this.UitstroomBlad.UitstroomRijen);

            List<UitstroomRij> uitstroomDiplomaAso = this.UitstroomBlad.FilterOpASO(uitstroomDiploma);
            List<UitstroomRij> uitstroomDiplomaTso = this.UitstroomBlad.FilterOpTSO(uitstroomDiploma);
            List<UitstroomRij> uitstroomDiplomaBso = this.UitstroomBlad.FilterOpBSO(uitstroomDiploma);
            List<UitstroomRij> uitstroomDiplomaKso = this.UitstroomBlad.FilterOpKSO(uitstroomDiploma);
            List<UitstroomRij> uitstroomDiplomaBlnd = this.UitstroomBlad.FilterOpAndereSO(uitstroomDiploma);

            this.PowerPoint.ChangeStudieduurSlide2(
                this.UitstroomBlad.FilterOpMinderDan3(uitstroomDiplomaAso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp3(uitstroomDiplomaAso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp4(uitstroomDiplomaAso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMeerDan4(uitstroomDiplomaAso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMinderDan3(uitstroomDiplomaTso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp3(uitstroomDiplomaTso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp4(uitstroomDiplomaTso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMeerDan4(uitstroomDiplomaTso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMinderDan3(uitstroomDiplomaBso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp3(uitstroomDiplomaBso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp4(uitstroomDiplomaBso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMeerDan4(uitstroomDiplomaBso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMinderDan3(uitstroomDiplomaKso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp3(uitstroomDiplomaKso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp4(uitstroomDiplomaKso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMeerDan4(uitstroomDiplomaKso, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMinderDan3(uitstroomDiplomaBlnd, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp3(uitstroomDiplomaBlnd, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOp4(uitstroomDiplomaBlnd, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)),
                this.UitstroomBlad.FilterOpMeerDan4(uitstroomDiplomaBlnd, EHBFunctions.GetCurrentAcademicYearBasedOnIndex(index)));
        }

        public void GenerateAll()
        {
            for (int i = 0; i < this.FilePathHandler.MaxAantalPaths; i++)
            {
                this.Load(i);

                List<InstroomRij> temp = this.GenerateInstroomData1(i + 1);
                this.GenerateInstroomData2(i + 1);
                this.GenerateInstroomData3(i + 1);

                this.GenerateDoorstroomData1(i + 1, temp);

                if (i == 0)
                {
                    this.GenerateDoorstroomData2();
                    this.GenerateStudieduur2(i + 1);
                }

                this.GenerateUitstroomData1(i + 1);
                this.GenerateStudieduur1(i + 1);
            }
        }

        public void SavePowerPoint()
        {
            this.PowerPoint.Save();
        }
    }
}
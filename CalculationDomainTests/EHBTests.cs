using Xunit;
using CalculationDomain.ErasmusHogeSchool.Instroom;
using CalculationDomain.ErasmusHogeSchool.Doorstroom;
using CalculationDomain.ErasmusHogeSchool.Uitstroom;
using CalculationDomain.ErasmusHogeSchool;
using System.Collections.Generic;

namespace CalculationDomainTests
{
    public class EHBTests
    {
        private const string Opleiding = "Bachelor in de Vroedkunde";
        private const string AbsPath = @"d:\GitHub\GIP\Documentation\";

        private const string InstroomPath = AbsPath + "Kopie van Instroom dd.08.10.2019 - 19-20.xlsx";
        private const string DoorstroomPath = AbsPath + "Kopie van Doorstroom dd.08.10.2019 -18-19.xlsx";
        private const string UitstroomPath = AbsPath + "Kopie van Uitstroom dd.08.10.2019 - 18-19.xlsx";

        // d:\GitHub\GIP\Documentation\
        // C:\Users\joren.schelkens.BAZANDPOORT.000\Documents\GitHub\GIP\Documentation\

        #region Instroom

        [Fact]
        public void LeesBestandInInstroomTest()
        {
            InstroomBlad instroomBlad = new InstroomBlad(InstroomPath, Opleiding);
            Assert.NotEmpty(instroomBlad.InstroomRijen);
        }

        [Fact]
        public void FilterOpVoltijdsEnNieuweStudentInstroomTest()
        {
            InstroomBlad instroomBlad = new InstroomBlad(InstroomPath, Opleiding);
            instroomBlad.FilterOpNieuweStudent();
            List<InstroomRij> temp = instroomBlad.FilterOpVoltijds();
            Assert.Equal(47, temp.Count);
        }

        #endregion

        #region Doorstroom

        [Fact]
        public void LeesBestandInDoorstroomTest()
        {
            DoorstroomBlad doorstroomBlad = new DoorstroomBlad(DoorstroomPath, Opleiding);
            Assert.NotEmpty(doorstroomBlad.DoorstroomRijen);
        }

        #endregion

        #region Uitstroom

        [Fact]
        public void LeesBestandInUitstroomTest()
        {
            UitstroomBlad uitstroomBlad = new UitstroomBlad(UitstroomPath, Opleiding);
            Assert.NotEmpty(uitstroomBlad.UitstroomRijen);
        }

        #endregion

        #region Main

        [Fact]
        public void MainLoadTest()
        {
            Main main = new Main(Opleiding);
            main.Load();

            Assert.NotNull(main.InstroomBlad);
            Assert.NotNull(main.DoorstroomBlad);
            Assert.NotNull(main.UitstroomBlad);
        }

        [Fact]
        public void MainPowerPointTest()
        {
            Main main = new Main(Opleiding);
            main.Load();

            main.GenerateInstroomData();
            main.SavePowerPoint();

            Assert.NotNull(main.PowerPoint);
        }

        #endregion

        #region PowerPoint

        [Fact]
        public void PowerPointSaveTest()
        {
            PowerPointClass powerPoint = new PowerPointClass(Opleiding);

            powerPoint.AddInstroomSlide(13, 12, 25, 10, 15, 20, 50);

            powerPoint.Save();

            Assert.NotNull(powerPoint);
        }

        #endregion
    }
}
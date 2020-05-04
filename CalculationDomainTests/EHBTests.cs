using CalculationDomain.ErasmusHogeSchool;
using CalculationDomain.ErasmusHogeSchool.Doorstroom;
using CalculationDomain.ErasmusHogeSchool.Instroom;
using CalculationDomain.ErasmusHogeSchool.Uitstroom;
using System.Collections.Generic;
using Xunit;

namespace CalculationDomainTests
{
    public class EHBTests
    {
        private const string Opleiding = "Bachelor in de Vroedkunde";
        private FilePathHandler FilePathHandler { get; set; } = new FilePathHandler();

        #region Instroom

        [Fact]
        public void LeesBestandInInstroomTest()
        {
            InstroomBlad instroomBlad = new InstroomBlad(this.FilePathHandler.InstroomPaths[0], Opleiding);
            Assert.NotEmpty(instroomBlad.InstroomRijen);
        }

        [Fact]
        public void FilterOpVoltijdsEnNieuweStudentInstroomTest()
        {
            InstroomBlad instroomBlad = new InstroomBlad(this.FilePathHandler.InstroomPaths[0], Opleiding);
            List<InstroomRij> instroomNieuweStudenten = instroomBlad.FilterOpNieuweStudent();
            List<InstroomRij> temp = instroomBlad.FilterOpVoltijds(instroomNieuweStudenten);

            Assert.Equal(47, temp.Count);
        }

        #endregion

        #region Doorstroom

        [Fact]
        public void LeesBestandInDoorstroomTest()
        {
            DoorstroomBlad doorstroomBlad = new DoorstroomBlad(this.FilePathHandler.DoorstroomPaths[0], Opleiding);
            Assert.NotEmpty(doorstroomBlad.DoorstroomRijen);
        }

        #endregion

        #region Uitstroom

        [Fact]
        public void LeesBestandInUitstroomTest()
        {
            UitstroomBlad uitstroomBlad = new UitstroomBlad(this.FilePathHandler.UitstroomPaths[0], Opleiding);
            Assert.NotEmpty(uitstroomBlad.UitstroomRijen);
        }

        #endregion

        #region Main

        [Fact]
        public void MainPowerPointTest()
        {
            Main main = new Main(Opleiding);
            main.GenerateAll();

            main.SavePowerPoint();

            Assert.NotNull(main.PowerPoint);
        }

        #endregion

        #region PowerPoint

        [Fact]
        public void PowerPointSaveTest()
        {
            PowerPointClass powerPoint = new PowerPointClass(Opleiding);

            powerPoint.Save();

            Assert.NotNull(powerPoint);
        }

        [Fact]
        public void PowerPointOpenTest()
        {
            PowerPointClass powerPoint = new PowerPointClass(Opleiding);

            Assert.NotNull(powerPoint);
        }

        [Fact]
        public void PowerPointTestMethodTest()
        {
            PowerPointClass powerPoint = new PowerPointClass(Opleiding);

            powerPoint.TestMethod();

            powerPoint.Save();

            Assert.NotNull(powerPoint);
        }

        #endregion

        #region Others

        [Fact]
        public void FilePathHandlerTest()
        {
            FilePathHandler file = new FilePathHandler();
            Assert.NotNull(file.InstroomPaths);
        }

        #endregion
    }
}
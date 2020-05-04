using CalculationDomain.ErasmusHogeSchool;
using DefaultDomain.ExcelReading;
using Xunit;

namespace DefaultDomainTests
{
    public class ExcelReadTests
    {
        private FilePathHandler FilePathHandler { get; set; } = new FilePathHandler();

        [Fact]
        public void ReadEHBTest()
        {
            Assert.NotEmpty(ExcelRead.ReadEHB(this.FilePathHandler.InstroomPaths[0]));
            Assert.NotEmpty(ExcelRead.ReadEHB(this.FilePathHandler.DoorstroomPaths[0]));
            Assert.NotEmpty(ExcelRead.ReadEHB(this.FilePathHandler.UitstroomPaths[0]));
        }
    }
}
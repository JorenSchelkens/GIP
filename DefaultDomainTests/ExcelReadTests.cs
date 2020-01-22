using Xunit;
using DefaultDomain.ExcelReading;
using CalculationDomain.ErasmusHogeSchool;

namespace DefaultDomainTests
{
    public class ExcelReadTests
    {
        private FilePathHandler FilePathHandler { get; set; } = new FilePathHandler();

        [Fact]
        public void ReadEHBTest()
        {
            Assert.NotEmpty(ExcelRead.ReadEHB(FilePathHandler.InstroomPaths[0]));
            Assert.NotEmpty(ExcelRead.ReadEHB(FilePathHandler.DoorstroomPaths[0]));
            Assert.NotEmpty(ExcelRead.ReadEHB(FilePathHandler.UitstroomPaths[0]));
        }
    }
}
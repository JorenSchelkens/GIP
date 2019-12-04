using Xunit;
using DefaultDomain.ExcelReading;

namespace DefaultDomainTests
{
    public class ExcelReadTests
    {
        private const string AbsPath = @"d:\GitHub\GIP\Documentation\";

        private const string InstroomPath = AbsPath + "Kopie van Instroom dd.08.10.2019 - 19-20.xlsx";
        private const string DoorstroomPath = AbsPath + "Kopie van Doorstroom dd.08.10.2019 -18-19.xlsx";
        private const string UitstroomPath = AbsPath + "Kopie van Uitstroom dd.08.10.2019 - 18-19.xlsx";

        //d:\GitHub\GIP\Documentation\
        //C:\Users\joren.schelkens.BAZANDPOORT.000\Documents\GitHub\GIP\Documentation\

        [Fact]
        public void ReadEHBTest()
        {
            Assert.NotEmpty(ExcelRead.ReadEHB(InstroomPath));
            Assert.NotEmpty(ExcelRead.ReadEHB(DoorstroomPath));
            Assert.NotEmpty(ExcelRead.ReadEHB(UitstroomPath));
        }
    }
}
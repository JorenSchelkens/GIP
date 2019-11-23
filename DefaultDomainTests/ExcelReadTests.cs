using Xunit;
using DefaultDomain.ExcelReading;

namespace DefaultDomainTests
{
    public class ExcelReadTests
    {
        [Fact]
        public void ReadEHBTest()
        {
            Assert.NotEmpty(ExcelRead.ReadEHB(@"d:\GitHub\GIP\Documentation\Kopie van Instroom dd.08.10.2019 - 19-20.xlsx"));
            Assert.NotEmpty(ExcelRead.ReadEHB(@"d:\GitHub\GIP\Documentation\Kopie van Doorstroom dd.08.10.2019 -18-19.xlsx"));
            Assert.NotEmpty(ExcelRead.ReadEHB(@"d:\GitHub\GIP\Documentation\Kopie van Uitstroom dd.08.10.2019 - 18-19.xlsx"));
        }
    }
}
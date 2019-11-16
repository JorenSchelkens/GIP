using Xunit;
using CalculationDomain.ErasmusHogeSchool.Instroom;

namespace CalculationDomainTests
{
    public class EHBTests
    {
        [Fact]
        public void InstroomTest()
        {
            InstroomBlad instroomBlad = 
                new InstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Instroom dd.08.10.2019 - 19-20.xlsx", 
                "Bachelor in de Vroedkunde");
            Assert.NotEmpty(instroomBlad.instroomRijen);
        }
    }
}

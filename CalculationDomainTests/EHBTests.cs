using Xunit;
using CalculationDomain.ErasmusHogeSchool.Instroom;
using CalculationDomain.ErasmusHogeSchool.Doorstroom;
using CalculationDomain.ErasmusHogeSchool.Uitstroom;

namespace CalculationDomainTests
{
    public class EHBTests
    {
        [Fact]
        public void LeesBestandInInstroomTest()
        {
            InstroomBlad instroomBlad = 
                new InstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Instroom dd.08.10.2019 - 19-20.xlsx", 
                "Bachelor in de Vroedkunde");
            Assert.NotEmpty(instroomBlad.InstroomRijen);
        }

        [Fact]
        public void LeesBestandInDoorstroomTest()
        {
            DoorstroomBlad doorstroomBlad =
                new DoorstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Doorstroom dd.08.10.2019 -18-19.xlsx",
                "Bachelor in de Vroedkunde");
            Assert.NotEmpty(doorstroomBlad.DoorstroomRijen);
        }

        [Fact]
        public void LeesBestandInUitstroomTest()
        {
            UitstroomBlad uitstroomBlad =
                new UitstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Uitstroom dd.08.10.2019 - 18-19.xlsx",
                "Bachelor in de Vroedkunde");
            Assert.NotEmpty(uitstroomBlad.UitstroomRijen);
        }
    }
}
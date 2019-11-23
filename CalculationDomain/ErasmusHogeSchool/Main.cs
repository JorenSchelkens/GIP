using CalculationDomain.ErasmusHogeSchool.Instroom;
using CalculationDomain.ErasmusHogeSchool.Doorstroom;
using CalculationDomain.ErasmusHogeSchool.Uitstroom;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class Main
    {
        private InstroomBlad InstroomBlad { get; set; }
        private DoorstroomBlad DoorstroomBlad { get; set; }
        private UitstroomBlad UitstroomBlad { get; set; }

        private string Filter { get; set; }

        public Main(string opleiding)
        {
            this.Filter = opleiding;
        }

        public void Load()
        {
            //Haal bestanden
            this.InstroomBlad = new InstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Instroom dd.08.10.2019 - 19-20.xlsx", this.Filter);
            this.DoorstroomBlad = new DoorstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Doorstroom dd.08.10.2019 -18-19.xlsx", this.Filter);
        }

    }
}
using CalculationDomain.ErasmusHogeSchool.Instroom;
using CalculationDomain.ErasmusHogeSchool.Doorstroom;
using CalculationDomain.ErasmusHogeSchool.Uitstroom;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class Main
    {
        public InstroomBlad InstroomBlad { get; set; }
        public DoorstroomBlad DoorstroomBlad { get; set; }
        public UitstroomBlad UitstroomBlad { get; set; }
        private string Filter { get; set; }

        public Main(string opleiding)
        {
            this.Filter = opleiding;
        }

        public void Load()
        {
            //Haal bestanden

            // d:\GitHub\GIP\Documentation\
            // C:\Users\joren.schelkens.BAZANDPOORT.000\Documents\GitHub\GIP\Documentation\

            this.InstroomBlad = new InstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Instroom dd.08.10.2019 - 19-20.xlsx", this.Filter);
            this.DoorstroomBlad = new DoorstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Doorstroom dd.08.10.2019 -18-19.xlsx", this.Filter);
            this.UitstroomBlad = new UitstroomBlad(@"d:\GitHub\GIP\Documentation\Kopie van Uitstroom dd.08.10.2019 - 18-19.xlsx", this.Filter);
        }

    }
}
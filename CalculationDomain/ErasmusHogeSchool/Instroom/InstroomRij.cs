namespace CalculationDomain.ErasmusHogeSchool.Instroom
{
    public class InstroomRij
    {
        public bool NieuweStudent { get; set; }
        public bool GeneratieStudent { get; set; }
        public string SoOnderwijsvorm { get; set; }
        public int Trajectschijfverdeling { get; set; }

        public override string ToString()
        {
            return "Trajectschijfverdeling: " + this.Trajectschijfverdeling;
        }
    }
}
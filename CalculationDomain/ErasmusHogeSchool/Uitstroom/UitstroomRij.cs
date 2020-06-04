namespace CalculationDomain.ErasmusHogeSchool.Uitstroom
{
    public class UitstroomRij
    {
        public string SoOnderwijsvorm { get; set; }
        public string Stamnummer { get; set; }
        public bool DiplomaBehaald { get; set; }

        public override string ToString()
        {
            return $"{this.SoOnderwijsvorm} , {this.Stamnummer}";
        }
    }
}
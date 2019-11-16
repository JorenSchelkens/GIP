namespace CalculationDomain
{
    public class DoorstroomRij
    {
        public int StudiepuntenTeVolgen { get; set; }
        public int StudiepuntenCredits { get; set; }
        public bool VolgtOlodInSchijf1 { get; set; }
        public bool NieuweStudentInInstelling { get; set; }
        //Trajectschijfverdeling => datatype?
        public string SoOnderwijsvorm { get; set; }
        public string Stamnummer { get; set; }
        public string KanDiplomaBehalen { get; set; }
        public bool HeeftDiplomaBehaalt { get; set; }
        public bool Generatie { get; set; }

    }
}
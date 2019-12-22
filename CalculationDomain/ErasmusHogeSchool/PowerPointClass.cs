using Syncfusion.Presentation;
using System;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class PowerPointClass
    {
        //https://help.syncfusion.com/file-formats/presentation/getting-started
        //https://www.asknumbers.com/centimeters-to-points.aspx
        public IPresentation PowerPoint { get; set; } = Presentation.Create();

        public PowerPointClass()
        {
            AddInstroomIntroSlide();
            AddInstroomSlide();

            Save();
        }

        public void AddInstroomIntroSlide()
        {
            ISlide slide = this.PowerPoint.Slides.Add(SlideLayoutType.Blank);
            IShape textShape = slide.AddTextBox(100, 75, 756, 100);
            IParagraph paragraph = textShape.TextBody.AddParagraph();
            paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            ITextPart textPart = paragraph.AddTextPart("1. Instroom");

            textPart.Font.FontSize = 40;
            textPart.Font.Bold = true;
        }

        public void AddInstroomSlide()
        {
            ISlide slide = this.PowerPoint.Slides.Add(SlideLayoutType.Blank);

            ITable table = slide.Tables.AddTable(8, 6, 100, 75, 583.94, 352.35);

            DateTime currentYearTemp = DateTime.Today;
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = DateTime.Today.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }

            table.Columns[0].Cells[1].TextBody.Text = "Voltijds";
            table.Columns[0].Cells[2].TextBody.Text = "Deeltijds";
            table.Columns[0].Cells[3].TextBody.Text = "Totaal";
            table.Columns[0].Cells[4].TextBody.Text = "Generatiestudent";
            table.Columns[0].Cells[5].TextBody.Text = "Niet-generatiestudent";
            table.Columns[0].Cells[6].TextBody.Text = "Aandeel generatiestud. in totale instroom";
            table.Columns[0].Cells[7].TextBody.Text = "Aandeel generatiestud. in voltijdse instroom";
        }

        public void Save()
        {
            PowerPoint.Save("Cijferanalyse.pptx");
            PowerPoint.Close();
        }
    }
}
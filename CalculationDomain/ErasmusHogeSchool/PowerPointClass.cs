using Syncfusion.Presentation;
using System;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class PowerPointClass
    {
        //https://help.syncfusion.com/file-formats/presentation/getting-started
        //https://www.asknumbers.com/centimeters-to-points.aspx
        public IPresentation PowerPoint { get; set; } = Presentation.Create();
        private string Opleiding { get; set; }

        public PowerPointClass(string opleiding)
        {
            this.Opleiding = opleiding;
            this.AddFirstSlide(this.Opleiding);
            this.AddInstroomIntroSlide1();
        }

        public void AddFirstSlide(string opleiding)
        {
            ISlide slide = this.PowerPoint.Slides.Add(SlideLayoutType.Blank);
            IShape textShape = slide.AddTextBox(100, 75, 756, 100);
            IParagraph paragraph = textShape.TextBody.AddParagraph();
            paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            ITextPart textPart = paragraph.AddTextPart(opleiding);

            textPart.Font.FontSize = 40;
            textPart.Font.Bold = true;
        }

        public void AddInstroomIntroSlide1()
        {
            ISlide slide = this.PowerPoint.Slides.Add(SlideLayoutType.Blank);
            IShape textShape = slide.AddTextBox(100, 75, 756, 100);
            IParagraph paragraph = textShape.TextBody.AddParagraph();
            paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            ITextPart textPart = paragraph.AddTextPart("1. Instroom");

            textPart.Font.FontSize = 40;
            textPart.Font.Bold = true;
        }

        public void AddInstroomSlide1(
            int voltijds, 
            int deeltijds, 
            int totaal, 
            int generatieStudent, 
            int nietGeneratieStudent, 
            int aandelInTotaal,
            int aandeelInVoltijds)
        {
            ISlide slide = this.PowerPoint.Slides.Add(SlideLayoutType.Blank);

            IShape textShape = slide.AddTextBox(100, 0, 756, 100);
            IParagraph paragraph = textShape.TextBody.AddParagraph();
            paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            ITextPart textPart = paragraph.AddTextPart("1.1 Instroom");

            ITable table = slide.Tables.AddTable(8, 6, 50, 40, 814.96, 435.57);

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

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = voltijds.ToString();
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = deeltijds.ToString();
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = totaal.ToString();
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = generatieStudent.ToString();
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = nietGeneratieStudent.ToString();
            table.Columns[table.Columns.Count - 1].Cells[6].TextBody.Text = $"{aandelInTotaal.ToString()}%";
            table.Columns[table.Columns.Count - 1].Cells[7].TextBody.Text = $"{aandeelInVoltijds.ToString()}%";
        }

        public void AddInstroomSlide2(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int totaal)
        {
            ISlide slide = this.PowerPoint.Slides.Add(SlideLayoutType.Blank);

            IShape textShape = slide.AddTextBox(100, 0, 756, 100);
            IParagraph paragraph = textShape.TextBody.AddParagraph();
            paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            ITextPart textPart = paragraph.AddTextPart("1.2 Instroom - type SO");

            ITable table = slide.Tables.AddTable(7, 9, 50, 40, 814.96, 435.57);

            table.Rows[0].Cells[0].TextBody.Text = "Type SO";

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

            table.Columns[0].Cells[1].TextBody.Text = "ASO";
            table.Columns[0].Cells[2].TextBody.Text = "TSO";
            table.Columns[0].Cells[3].TextBody.Text = "BSO";
            table.Columns[0].Cells[4].TextBody.Text = "KSO";
            table.Columns[0].Cells[5].TextBody.Text = "Buitenland of geen info";
            table.Columns[0].Cells[6].TextBody.Text = "TOT";

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = FormatString1(aso);
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = FormatString1(tso);
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = FormatString1(bso);
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = FormatString1(kso);
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = FormatString1(buiteland);
            table.Columns[table.Columns.Count - 1].Cells[6].TextBody.Text = FormatString1(totaal);
        }

        public void AddInstroomSlide3(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland)
        {
            ISlide slide = this.PowerPoint.Slides.Add(SlideLayoutType.Blank);

            IShape textShape = slide.AddTextBox(100, 0, 756, 100);
            IParagraph paragraph = textShape.TextBody.AddParagraph();
            paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            ITextPart textPart = paragraph.AddTextPart("1.3 Instroom - type SO");

            ITable table = slide.Tables.AddTable(6, 9, 50, 40, 814.96, 435.57);

            table.Rows[0].Cells[0].TextBody.Text = "Type SO";

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

            table.Columns[0].Cells[1].TextBody.Text = "ASO";
            table.Columns[0].Cells[2].TextBody.Text = "TSO";
            table.Columns[0].Cells[3].TextBody.Text = "BSO";
            table.Columns[0].Cells[4].TextBody.Text = "KSO";
            table.Columns[0].Cells[5].TextBody.Text = "Buitenland of geen info";

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = FormatString2(aso);
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = FormatString2(tso);
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = FormatString2(bso);
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = FormatString2(kso);
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = FormatString2(buiteland);
        }

        public void Save()
        {
            PowerPoint.Save($"{this.Opleiding}.pptx");
            PowerPoint.Close();
        }

        private string FormatString1(int toFormat)
        {
            string temp;

            if (toFormat == 0)
            {
                temp = "-";
            }
            else
            {
                temp = toFormat.ToString();
            }

            return temp;
        }

        private string FormatString2(int toFormat)
        {
            string temp;

            if (toFormat == 0)
            {
                temp = "-";
            }
            else
            {
                temp = $"{toFormat.ToString()}%";
            }

            return temp;
        }
    }
}
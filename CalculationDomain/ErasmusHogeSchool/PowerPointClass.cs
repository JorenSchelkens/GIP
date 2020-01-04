﻿using Syncfusion.Presentation;
using System;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class PowerPointClass
    {
        //https://help.syncfusion.com/file-formats/presentation/getting-started
        //https://www.asknumbers.com/centimeters-to-points.aspx
        //https://help.syncfusion.com/file-formats/presentation/working-with-tables#modifying-the-table

        // D:\GitHub\GIP\CalculationDomain\ErasmusHogeSchool\EmptyPowerPoint.pptx
        // C:\Users\joren.schelkens.BAZANDPOORT.000\Documents\GitHub\GIP\CalculationDomain\ErasmusHogeSchool\EmptyPowerPoint.pptx

        public IPresentation PowerPoint { get; set; } = Presentation.Open(@"D:\GitHub\GIP\CalculationDomain\ErasmusHogeSchool\EmptyPowerPoint.pptx");
        private string Opleiding { get; set; }

        public PowerPointClass(string opleiding)
        {
            this.Opleiding = opleiding;

            this.ChangeFirstSlide();
        }

        public void TestMethod()
        {
            ChangeDoorstroomSlide1();
        }

        public void ChangeFirstSlide()
        {
            ISlide slide = this.PowerPoint.Slides[0];
            IShape shape = slide.Shapes[0] as IShape;
            IParagraph paragraph = shape.TextBody.Paragraphs[0];
            ITextPart textPart = paragraph.TextParts[0];

            textPart.Text = this.Opleiding;
        }

        public void ChangeInstroomSlide1(
            int voltijds,
            int deeltijds,
            int totaal,
            int generatieStudent,
            int nietGeneratieStudent,
            int aandelInTotaal,
            int aandeelInVoltijds)
        {

            ISlide slide = this.PowerPoint.Slides[5];
            ITable table = slide.Tables[0];

            DateTime currentYearTemp = GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = FormatStringNonPercent(voltijds);
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = FormatStringNonPercent(deeltijds);
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = FormatStringNonPercent(totaal);
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = FormatStringNonPercent(generatieStudent);
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = FormatStringNonPercent(nietGeneratieStudent);
            table.Columns[table.Columns.Count - 1].Cells[6].TextBody.Text = FormatStringPercent(aandelInTotaal);
            table.Columns[table.Columns.Count - 1].Cells[7].TextBody.Text = FormatStringPercent(aandeelInVoltijds);
        }

        public void ChangeInstroomSlide2(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int totaal)
        {
            ISlide slide = this.PowerPoint.Slides[6];
            ITable table = slide.Tables[0];

            DateTime currentYearTemp = GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = FormatStringNonPercent(aso);
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = FormatStringNonPercent(tso);
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = FormatStringNonPercent(bso);
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = FormatStringNonPercent(kso);
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = FormatStringNonPercent(buiteland);
            table.Columns[table.Columns.Count - 1].Cells[7].TextBody.Text = FormatStringNonPercent(totaal);
        }

        public void ChangeInstroomSlide3(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland)
        {
            ISlide slide = this.PowerPoint.Slides[7];
            ITable table = slide.Tables[0];

            DateTime currentYearTemp = GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = FormatStringPercent(aso);
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = FormatStringPercent(tso);
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = FormatStringPercent(bso);
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = FormatStringPercent(kso);
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = FormatStringPercent(buiteland);
        }

        public void ChangeInstroomSlide4(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int aantal)
        {
            ISlide slide = this.PowerPoint.Slides[8];
            ITable table = slide.Tables[0];

            DateTime currentYearTemp = GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = FormatStringNonPercent(aso);
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = FormatStringNonPercent(tso);
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = FormatStringNonPercent(bso);
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = FormatStringNonPercent(kso);
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = FormatStringNonPercent(buiteland);
            table.Columns[table.Columns.Count - 1].Cells[7].TextBody.Text = FormatStringNonPercent(aantal);
        }

        public void ChangeInstroomSlide5(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland)
        {
            ISlide slide = this.PowerPoint.Slides[9];
            ITable table = slide.Tables[0];

            DateTime currentYearTemp = GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }

            table.Columns[table.Columns.Count - 1].Cells[1].TextBody.Text = FormatStringPercent(aso);
            table.Columns[table.Columns.Count - 1].Cells[2].TextBody.Text = FormatStringPercent(tso);
            table.Columns[table.Columns.Count - 1].Cells[3].TextBody.Text = FormatStringPercent(bso);
            table.Columns[table.Columns.Count - 1].Cells[4].TextBody.Text = FormatStringPercent(kso);
            table.Columns[table.Columns.Count - 1].Cells[5].TextBody.Text = FormatStringPercent(buiteland);
        }

        public void ChangeDoorstroomSlide1()
        {
            ISlide slide = this.PowerPoint.Slides[12];
            ITable table = slide.Tables[0];

            DateTime currentYearTemp = GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }



            //TABEL 2 --> rechts boven
            table = slide.Tables[1];

            currentYearTemp = GetCurrentAcademicYear();
            currentYear = currentYearTemp.Year.ToString();
            nextYearTemp = currentYearTemp.AddYears(1);
            nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = $"´{currentYear.Substring(2)}-´{nextYear.Substring(2)}";

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }
        }

        public void Save()
        {
            PowerPoint.Save($"{this.Opleiding}.pptx");
            PowerPoint.Close();
        }

        private string FormatStringNonPercent(int toFormat)
        {
            if (toFormat == 0)
            {
                return "-";
            }
            return toFormat.ToString();
        }

        private string FormatStringPercent(int toFormat)
        {
            if (toFormat == 0)
            {
                return "-";
            }
            return $"{toFormat.ToString()}%";
        }

        private DateTime GetCurrentAcademicYear()
        {
            DateTime current = DateTime.Now;

            if (DateTime.Now.Month < 9)
            {
                current = current.AddYears(-1);
            }

            return current;
        }
    }
}
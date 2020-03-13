using Syncfusion.Presentation;
using System;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class PowerPointClass
    {
        //https://help.syncfusion.com/file-formats/presentation/getting-started
        //https://www.asknumbers.com/centimeters-to-points.aspx
        //https://help.syncfusion.com/file-formats/presentation/working-with-tables#modifying-the-table

        public IPresentation PowerPoint { get; set; }
        private string Opleiding { get; set; }
        private FilePathHandler FilePathHandler { get; set; } = new FilePathHandler();

        public PowerPointClass(string opleiding)
        {
            this.PowerPoint = Presentation.Open(this.FilePathHandler.PowerPointPath);
            this.Opleiding = opleiding;

            this.ChangeFirstSlide();
        }

        public void TestMethod()
        {
            ChangeUitstroomSlide1(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1);
        }

        private void ChangeTableHeading(ITable table)
        {
            DateTime currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = EHBFunctions.FormatYearString(currentYear, nextYear);

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }
        }

        private void ChangeTableHeadingSpecial(ITable table)
        {
            DateTime currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Rows[0].Cells.Count - 1; i > 0; i--)
            {
                table.Rows[0].Cells[i].TextBody.Text = EHBFunctions.FormatYearStringSpecial(currentYear, nextYear);

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }
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
            int aandeelInVoltijds,
            int index)
        {

            ISlide slide = this.PowerPoint.Slides[5];
            ITable table = slide.Tables[0];

            if (index == 1)
            {
                ChangeTableHeading(table);
            }

            table.Columns[table.Columns.Count - index].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(voltijds);
            table.Columns[table.Columns.Count - index].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(deeltijds);
            table.Columns[table.Columns.Count - index].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(totaal);
            table.Columns[table.Columns.Count - index].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(generatieStudent);
            table.Columns[table.Columns.Count - index].Cells[5].TextBody.Text = EHBFunctions.FormatStringNonPercent(nietGeneratieStudent);
            table.Columns[table.Columns.Count - index].Cells[6].TextBody.Text = EHBFunctions.FormatStringPercent(aandelInTotaal);
            table.Columns[table.Columns.Count - index].Cells[7].TextBody.Text = EHBFunctions.FormatStringPercent(aandeelInVoltijds);
        }

        public void ChangeInstroomSlide2(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int totaal,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[6];
            ITable table = slide.Tables[0];

            if (index == 1)
            {
                ChangeTableHeading(table);
            }

            table.Columns[table.Columns.Count - index].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(aso);
            table.Columns[table.Columns.Count - index].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(tso);
            table.Columns[table.Columns.Count - index].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(bso);
            table.Columns[table.Columns.Count - index].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(kso);
            table.Columns[table.Columns.Count - index].Cells[5].TextBody.Text = EHBFunctions.FormatStringNonPercent(buiteland);
            table.Columns[table.Columns.Count - index].Cells[7].TextBody.Text = EHBFunctions.FormatStringNonPercent(totaal);
        }

        public void ChangeInstroomSlide3(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[7];
            ITable table = slide.Tables[0];

            if (index == 1)
            {
                ChangeTableHeading(table);
            }

            table.Columns[table.Columns.Count - index].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(aso);
            table.Columns[table.Columns.Count - index].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tso);
            table.Columns[table.Columns.Count - index].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(bso);
            table.Columns[table.Columns.Count - index].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(kso);
            table.Columns[table.Columns.Count - index].Cells[5].TextBody.Text = EHBFunctions.FormatStringPercent(buiteland);
        }

        public void ChangeInstroomSlide4(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int aantal,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[8];
            ITable table = slide.Tables[0];

            if (index == 1)
            {
                ChangeTableHeading(table);
            }

            table.Columns[table.Columns.Count - index].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(aso);
            table.Columns[table.Columns.Count - index].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(tso);
            table.Columns[table.Columns.Count - index].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(bso);
            table.Columns[table.Columns.Count - index].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(kso);
            table.Columns[table.Columns.Count - index].Cells[5].TextBody.Text = EHBFunctions.FormatStringNonPercent(buiteland);
            table.Columns[table.Columns.Count - index].Cells[7].TextBody.Text = EHBFunctions.FormatStringNonPercent(aantal);
        }

        public void ChangeInstroomSlide5(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[9];
            ITable table = slide.Tables[0];

            if (index == 1)
            {
                ChangeTableHeading(table);
            }

            table.Columns[table.Columns.Count - index].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(aso);
            table.Columns[table.Columns.Count - index].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tso);
            table.Columns[table.Columns.Count - index].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(bso);
            table.Columns[table.Columns.Count - index].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(kso);
            table.Columns[table.Columns.Count - index].Cells[5].TextBody.Text = EHBFunctions.FormatStringPercent(buiteland);
        }

        public void ChangeDoorstroomSlide1(
            int zestigStp1,
            int tussenZestigStpEnVijfenveertig1,
            int onderVijfenveertig1,
            int dropOut,
            int zestigStp2,
            int tussenZestigStpEnVijfenveertig2,
            int onderVijfenveertig2,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[12];
            ITable table = slide.Tables[0];

            IShape item = slide.Shapes[1] as IShape;
            item.TextBody.Paragraphs[0].Text = "Dit is een test 1";

            IShape item2 = slide.Shapes[6] as IShape;
            item2.TextBody.Paragraphs[0].Text = "Dit is een test 2";

            IShape item3 = slide.Shapes[8] as IShape;
            item3.TextBody.Paragraphs[0].Text = "Dit is een test 3";

            if (table.Columns.Count - index - 1 != 0)
            {
                if (index == 1)
                {
                    ChangeTableHeadingSpecial(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(zestigStp2);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tussenZestigStpEnVijfenveertig2);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(onderVijfenveertig2);

                //TABEL 2 --> rechts boven
                table = slide.Tables[1];

                if (index == 1)
                {
                    ChangeTableHeading(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(zestigStp1);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tussenZestigStpEnVijfenveertig1);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(onderVijfenveertig1);
                table.Columns[table.Columns.Count - index - 1].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(dropOut);
            }
        }

        public void ChangeUitstroomSlide1(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int aantal,
            int asoP,
            int tsoP,
            int bsoP,
            int ksoP,
            int buitelandP,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[16];
            ITable table = slide.Tables[0];

            if (table.Columns.Count - index - 1 != 0)
            {
                if (index == 1)
                {
                    ChangeTableHeading(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(aso);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(tso);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(bso);
                table.Columns[table.Columns.Count - index - 1].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(kso);
                table.Columns[table.Columns.Count - index - 1].Cells[5].TextBody.Text = EHBFunctions.FormatStringNonPercent(buiteland);
                table.Columns[table.Columns.Count - index - 1].Cells[7].TextBody.Text = EHBFunctions.FormatStringNonPercent(aantal);

                //TABEL 2 --> Onderaan
                table = slide.Tables[1];

                if (index == 1)
                {
                    ChangeTableHeading(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(asoP);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tsoP);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(bsoP);
                table.Columns[table.Columns.Count - index - 1].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(ksoP);
                table.Columns[table.Columns.Count - index - 1].Cells[5].TextBody.Text = EHBFunctions.FormatStringPercent(buitelandP);
            }
        }

        public void Save()
        {
            PowerPoint.Save($"Cijferanalyse {this.Opleiding}.pptx");
            PowerPoint.Close();
        }
    }
}
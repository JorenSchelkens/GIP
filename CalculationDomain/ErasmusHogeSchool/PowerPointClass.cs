using Syncfusion.Presentation;
using System;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class PowerPointClass
    {
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
            this.ChangeStudierendementSlide1(1, 1, 0);
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

        private void ChangeTableHeading2(ITable table)
        {
            DateTime currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime nextYearTemp = currentYearTemp.AddYears(1);
            string nextYear = nextYearTemp.Year.ToString();

            for (int i = table.Columns[0].Cells.Count - 1; i > 0; i--)
            {
                table.Columns[0].Cells[i].TextBody.Text = EHBFunctions.FormatYearString(currentYear, nextYear);

                currentYearTemp = currentYearTemp.AddYears(-1);
                currentYear = currentYearTemp.Year.ToString();
                nextYearTemp = nextYearTemp.AddYears(-1);
                nextYear = nextYearTemp.Year.ToString();
            }
        }

        private void ChangeTableHeading3(ITable table)
        {
            DateTime currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            currentYearTemp = currentYearTemp.AddYears(-1);

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

            shape = slide.Shapes[1] as IShape;
            shape.TextBody.Text += $"\r{EHBFunctions.GetCurrentAcademicYear().Year}";
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
                this.ChangeTableHeading(table);
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
                this.ChangeTableHeading(table);
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
                this.ChangeTableHeading(table);
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
                this.ChangeTableHeading(table);
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
                this.ChangeTableHeading(table);
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

            if (table.Columns.Count - index - 1 != 0)
            {
                if (index == 1)
                {
                    this.ChangeTableHeadingSpecial(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(zestigStp2);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tussenZestigStpEnVijfenveertig2);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(onderVijfenveertig2);

                IShape text;

                switch (index)
                {
                    case 1:
                        text = slide.Shapes[22] as IShape;
                        break;
                    case 2:
                        text = slide.Shapes[6] as IShape;
                        break;
                    case 3:
                        text = slide.Shapes[8] as IShape;
                        break;
                    case 4:
                        text = slide.Shapes[1] as IShape;
                        break;
                    default:
                        text = null;
                        break;
                }

                text.TextBody.Paragraphs[0].Text = EHBFunctions.FormatStringPercent(zestigStp2 + tussenZestigStpEnVijfenveertig2);

                //TABEL 2 --> rechts boven
                table = slide.Tables[1];

                if (index == 1)
                {
                    this.ChangeTableHeading(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(zestigStp1);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tussenZestigStpEnVijfenveertig1);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(onderVijfenveertig1);
                table.Columns[table.Columns.Count - index - 1].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(dropOut);

                switch (index)
                {
                    case 1:
                        text = slide.Shapes[20] as IShape;
                        break;
                    case 2:
                        text = slide.Shapes[12] as IShape;
                        break;
                    case 3:
                        text = slide.Shapes[14] as IShape;
                        break;
                    case 4:
                        text = slide.Shapes[16] as IShape;
                        break;
                    default:
                        text = null;
                        break;
                }

                text.TextBody.Paragraphs[0].Text = EHBFunctions.FormatStringPercent(zestigStp1 + tussenZestigStpEnVijfenveertig1);

                text = slide.Shapes[18] as IShape;
                text.TextBody.Paragraphs[0].Text = "";

                text = slide.Shapes[24] as IShape;
                text.TextBody.Paragraphs[0].Text = "";
            }
        }

        public void ChangeDoorstroomSlide2(
            int aso60,
            int asoTussen45En60,
            int aso45,
            int asoTotaal,
            int tso60,
            int tsoTussen45En60,
            int tso45,
            int tsoTotaal,
            int kso60,
            int ksoTussen45En60,
            int kso45,
            int ksoTotaal,
            int bso60,
            int bsoTussen45En60,
            int bso45,
            int bsoTotaal,
            int buiteland60,
            int buitelandTussen45En60,
            int buiteland45,
            int buitelandTotaal)
        {
            ISlide slide = this.PowerPoint.Slides[13];
            ITable table = slide.Tables[0];

            table.Columns[1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(aso60);
            table.Columns[1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(asoTussen45En60);
            table.Columns[1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(aso45);
            table.Columns[1].Cells[5].TextBody.Text = asoTotaal.ToString();

            table.Columns[2].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(tso60);
            table.Columns[2].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tsoTussen45En60);
            table.Columns[2].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(tso45);
            table.Columns[2].Cells[5].TextBody.Text = tsoTotaal.ToString();

            table.Columns[3].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(kso60);
            table.Columns[3].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(ksoTussen45En60);
            table.Columns[3].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(kso45);
            table.Columns[3].Cells[5].TextBody.Text = ksoTotaal.ToString();

            table.Columns[4].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(bso60);
            table.Columns[4].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(bsoTussen45En60);
            table.Columns[4].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(bso45);
            table.Columns[4].Cells[5].TextBody.Text = bsoTotaal.ToString();

            table.Columns[5].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(buiteland60);
            table.Columns[5].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(buitelandTussen45En60);
            table.Columns[5].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(buiteland45);
            table.Columns[5].Cells[5].TextBody.Text = buitelandTotaal.ToString();

            IShape text;

            text = slide.Shapes[6] as IShape;
            text.TextBody.Paragraphs[0].Text = EHBFunctions.FormatStringPercent(aso60 + asoTussen45En60);

            text = slide.Shapes[7] as IShape;
            text.TextBody.Paragraphs[0].Text = EHBFunctions.FormatStringPercent(tso60 + tsoTussen45En60);

            DateTime currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime previousYearTemp = currentYearTemp.AddYears(-1);
            string previousYear = previousYearTemp.Year.ToString();

            text = slide.Shapes[1] as IShape;
            text.TextBody.Text += " " + EHBFunctions.FormatYearStringSpecial(previousYear, currentYear);
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
                    this.ChangeTableHeading(table);
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
                    this.ChangeTableHeading(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(asoP);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tsoP);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(bsoP);
                table.Columns[table.Columns.Count - index - 1].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(ksoP);
                table.Columns[table.Columns.Count - index - 1].Cells[5].TextBody.Text = EHBFunctions.FormatStringPercent(buitelandP);
            }
        }

        public void ChangeStudieduurSlide1(
            int minderDanDrie,
            int drie,
            int vier,
            int meerDanVier,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[19];
            ITable table = slide.Tables[0];

            if (table.Rows.Count - index - 1 != 0)
            {
                if (index == 1)
                {
                    this.ChangeTableHeading2(table);
                }

                table.Rows[table.Rows.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(minderDanDrie);
                table.Rows[table.Rows.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(drie);
                table.Rows[table.Rows.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(vier);
                table.Rows[table.Rows.Count - index - 1].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(meerDanVier);
            }
        }

        public void ChangeStudieduurSlide2(
            int asoMinderDanDrie,
            int asoDrie,
            int asoVier,
            int asoMeerDanVier,
            int tsoMinderDanDrie,
            int tsoDrie,
            int tsoVier,
            int tsoMeerDanVier,
            int bsoMinderDanDrie,
            int bsoDrie,
            int bsoVier,
            int bsoMeerDanVier,
            int ksoMinderDanDrie,
            int ksoDrie,
            int ksoVier,
            int ksoMeerDanVier,
            int blndMinderDanDrie,
            int blndDrie,
            int blndVier,
            int blndMeerDanVier)
        {
            ISlide slide = this.PowerPoint.Slides[20];
            ITable table = slide.Tables[0];

            table.Rows[1].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoMinderDanDrie);
            table.Rows[1].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoDrie);
            table.Rows[1].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoVier);
            table.Rows[1].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoMeerDanVier);

            table.Rows[2].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(tsoMinderDanDrie);
            table.Rows[2].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(tsoDrie);
            table.Rows[2].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(tsoVier);
            table.Rows[2].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(tsoMeerDanVier);

            table.Rows[3].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(bsoMinderDanDrie);
            table.Rows[3].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(bsoDrie);
            table.Rows[3].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(bsoVier);
            table.Rows[3].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(bsoMeerDanVier);

            table.Rows[4].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(ksoMinderDanDrie);
            table.Rows[4].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(ksoDrie);
            table.Rows[4].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(ksoVier);
            table.Rows[4].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(ksoMeerDanVier);

            table.Rows[5].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(blndMinderDanDrie);
            table.Rows[5].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(blndDrie);
            table.Rows[5].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(blndVier);
            table.Rows[5].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(blndMeerDanVier);

            table.Rows[7].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoMinderDanDrie + tsoMinderDanDrie + bsoMinderDanDrie + ksoMinderDanDrie + blndMinderDanDrie);
            table.Rows[7].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoDrie + tsoDrie + bsoDrie + ksoDrie + blndDrie);
            table.Rows[7].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoVier + tsoVier + bsoVier + ksoVier + blndVier);
            table.Rows[7].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(asoMeerDanVier + tsoMeerDanVier + bsoMeerDanVier + ksoMeerDanVier + blndMeerDanVier);

            IShape text;

            DateTime currentYearTemp = EHBFunctions.GetCurrentAcademicYear();
            string currentYear = currentYearTemp.Year.ToString();
            DateTime previousYearTemp = currentYearTemp.AddYears(-1);
            string previousYear = previousYearTemp.Year.ToString();

            text = slide.Shapes[1] as IShape;
            text.TextBody.Text = $"4.2. Studieduur per type SO, uitstroom {EHBFunctions.FormatYearString(previousYear, currentYear)} \r in aantal studenten";
        }

        public void ChangeStudieduurSlide3(
            int aso,
            int tso,
            int bso,
            int kso,
            int buiteland,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[21];
            ITable table = slide.Tables[0];

            if (table.Columns.Count - index - 1 != 0)
            {
                if (index == 1)
                {
                    this.ChangeTableHeading(table);
                }

                table.Columns[table.Columns.Count - index - 1].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(aso);
                table.Columns[table.Columns.Count - index - 1].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(tso);
                table.Columns[table.Columns.Count - index - 1].Cells[3].TextBody.Text = EHBFunctions.FormatStringNonPercent(bso);
                table.Columns[table.Columns.Count - index - 1].Cells[4].TextBody.Text = EHBFunctions.FormatStringNonPercent(kso);
                table.Columns[table.Columns.Count - index - 1].Cells[5].TextBody.Text = EHBFunctions.FormatStringNonPercent(buiteland);
            }
        }

        public void ChangeStudierendementSlide1(
            int opgenomePunten,
            int verworvenPunten,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[24];
            ITable table = slide.Tables[0];
            //IPresentationChart chart = slide.Charts[0];

            if (index == 1)
            {
                this.ChangeTableHeading3(table);
            }

            table.Columns[table.Columns.Count - index].Cells[1].TextBody.Text = EHBFunctions.FormatStringNonPercent(opgenomePunten);
            table.Columns[table.Columns.Count - index].Cells[2].TextBody.Text = EHBFunctions.FormatStringNonPercent(verworvenPunten);
            table.Columns[table.Columns.Count - index].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent((int) Math.Round(((double)verworvenPunten / opgenomePunten) * 100.00));
        }

        public void ChangeStudierendementSlide2(
            int aso,
            int tso,
            int bso,
            int kso,
            int index)
        {
            ISlide slide = this.PowerPoint.Slides[25];
            ITable table = slide.Tables[0];
            //IPresentationChart chart = slide.Charts[0];

            if (index == 1)
            {
                this.ChangeTableHeading3(table);
            }

            table.Columns[table.Columns.Count - index].Cells[1].TextBody.Text = EHBFunctions.FormatStringPercent(aso);
            table.Columns[table.Columns.Count - index].Cells[2].TextBody.Text = EHBFunctions.FormatStringPercent(tso);
            table.Columns[table.Columns.Count - index].Cells[3].TextBody.Text = EHBFunctions.FormatStringPercent(bso);
            table.Columns[table.Columns.Count - index].Cells[4].TextBody.Text = EHBFunctions.FormatStringPercent(kso);
        }

        public void Save()
        {
            this.PowerPoint.Save($"Cijferanalyse {this.Opleiding}.pptx");
            this.PowerPoint.Close();
        }
    }
}
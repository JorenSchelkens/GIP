﻿using System.IO;
using CalculationDomain.ErasmusHogeSchool;

namespace GIP.Data
{
    public class PresentationService
    {
        public MemoryStream CreatePowerPoint(string opleiding)
        {
            Main main = new Main(opleiding);
            main.GenerateAll();

            //Save the PowerPoint Presentation as stream
            MemoryStream stream = new MemoryStream();
            main.PowerPoint.PowerPoint.Save(stream);

            //Close the PowerPoint Presentation as stream
            main.PowerPoint.PowerPoint.Close();
            stream.Position = 0;

            return stream;
        }
    }
}
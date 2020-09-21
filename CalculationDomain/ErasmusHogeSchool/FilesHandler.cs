using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class FilesHandler
    {
        public List<MemoryStream> DoorstroomExcels { get; set; } = new List<MemoryStream>();
        public List<MemoryStream> InstroomExcels { get; set; } = new List<MemoryStream>();
        public List<MemoryStream> UitstroomExcels { get; set; } = new List<MemoryStream>();
        public MemoryStream PowerPointPath { get; set; }
        public int MaxAantalPaths { get; set; } = 5;

        public FilesHandler(List<MemoryStream> excels)
        {
            this.SetFiles(excels);
        }

        private void SetFiles(List<MemoryStream> excels)
        {
            for (int i = excels.Count - 1; i >= 0; i--)
            {
                if(i == 15)
                {
                    this.PowerPointPath = excels[i];
                }
                else if(i < 15 && i > 9)
                {
                    this.UitstroomExcels.Add(excels[i]);
                }
                else if (i < 10 && i > 4)
                {
                    this.InstroomExcels.Add(excels[i]);
                }
                else if (i < 5)
                {
                    this.DoorstroomExcels.Add(excels[i]);
                }
            }
        }
    }
}
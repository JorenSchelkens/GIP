using System.Collections.Generic;
using System.IO;

namespace CalculationDomain.ErasmusHogeSchool
{
    public class FilesHandler
    {
        public List<byte[]> DoorstroomExcels { get; set; } = new List<byte[]>();
        public List<byte[]> InstroomExcels { get; set; } = new List<byte[]>();
        public List<byte[]> UitstroomExcels { get; set; } = new List<byte[]>();
        public byte[] PowerPointBytes { get; set; }
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
                    this.PowerPointBytes = excels[i].ToArray();
                }
                else if(i < 15 && i > 9)
                {
                    this.UitstroomExcels.Add(excels[i].ToArray());
                }
                else if (i < 10 && i > 4)
                {
                    this.InstroomExcels.Add(excels[i].ToArray());
                }
                else if (i < 5)
                {
                    this.DoorstroomExcels.Add(excels[i].ToArray());
                }
            }
        }

        public MemoryStream GetPowerPointStream()
        {
            return new MemoryStream(this.PowerPointBytes);
        }

        public MemoryStream GetInstroomStream(int index)
        {
            return new MemoryStream(this.InstroomExcels[index]);
        }

        public MemoryStream GetDoorstroomStream(int index)
        {
            return new MemoryStream(this.DoorstroomExcels[index]);
        }

        public MemoryStream GetUitstroomStream(int index)
        {
            return new MemoryStream(this.DoorstroomExcels[index]);
        }
    }
}
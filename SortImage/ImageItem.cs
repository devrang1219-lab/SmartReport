using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.SortImage
{
    public class ImageItem
    {
        public int Order { get; set; }

        public string FileName { get; set; }

        public string FullPath { get; set; }

        public DateTime LastWriteTime { get; set; }

        public override string ToString()
        {
            return FileName;
        }
    }
}

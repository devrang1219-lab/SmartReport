using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartReport
{
    public class SynologyFileGridItem
    {
        public bool Selected { get; set; }   // 체크박스용
        public string Name { get; set; }
        public string Path { get; set; }
        public long Size { get; set; }
        public string SizeText { get; set; }
        public DateTime ModifiedTime { get; set; }
        public string ModifiedTimeText { get; set; }

        public SynologyFileGridItem()
        {
            Name = "";
            Path = "";
            SizeText = "";
            ModifiedTimeText = "";
        }
    }
}

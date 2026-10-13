using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Comm
{
    public interface IReport
    {
        void AddLog(string part, string msg);
    }

}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ecodistrict.Excel
{
    public class ErrorMessage:EventArgs
    {
        public string SourceFunction;
        public string Message;
    }
}

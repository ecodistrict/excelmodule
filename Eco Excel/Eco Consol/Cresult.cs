using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eco_FormsTest
{
    public class Cresult
    {
        public Cresult()
        {
            Cells=new List<Cell>();
        }

        public string ParamName { get; set; }
        public string Sheetname { get; set; }
        public string ParamType { get; set; }
        public List<Cell> Cells { get; set; }
        public string DataType { get; set; }
    }
}

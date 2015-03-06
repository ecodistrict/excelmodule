using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eco_FormsTest;
using Microsoft.Office.Interop.Excel;

namespace EcoExcel
{
    public class CExcelRef
    {   public enum InOrOutParam
        {
            InParameter,
            OutParameter    
        }

        public CExcelRef()
        {
            Cells=new List<Cell>();
        }

        public string Sheetname { get; set; }
        public InOrOutParam ParamType { get; set; }
        public List<Cell> Cells { get; set; } 
        public Type DataType { get; set; }
    }
}

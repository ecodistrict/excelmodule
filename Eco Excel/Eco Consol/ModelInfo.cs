using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eco_Consol
{
    public class ModelInfo
    {
        public string ModuleName { get; set; }
        public string ModuleId { get; set; }
        public string ModuleDescription { get; set; }
        public List<string> KpiList { get; set; }
        
        public ModelInfo()
        {
            KpiList=new List<string>();
        }

    }
}

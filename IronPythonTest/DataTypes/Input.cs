using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTypes
{
    public class Input
    {
        protected string type { get; set; }
        protected string label { get; set; }
        protected string id { get; set; }

        public virtual string ToJason() { return ""; }
    }
}

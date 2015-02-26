using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTypes
{
    public class List : NonAtomic
    {

        public List(string label="", string id="")
        {
            this.type = "list";
            this.label = label;
            this.id = id;
        }

        public override void Add(Input item)
        {
            inputs.Add(item);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTypes
{
    public class Number : Atomic
    {
        decimal min { get; set; }
        decimal max { get; set; }
        decimal value { get; set; }

        public Number(string label="", string id ="", 
            decimal min=0, decimal max=10, decimal value=5)
        {
            this.type = "number";
            this.label = label;
            this.id = id;

            this.min = min;
            this.max = max;
            this.value = value;
        }


        public string ToJason()
        {

            return "";
        }
    }
}

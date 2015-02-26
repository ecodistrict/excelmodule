using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTypes
{
    public class InputSpecification : NonAtomic
    {
        public InputSpecification()
        {
            this.type = "";
            this.label = "";
            this.id = "";
        }

        public override void Add(Input item)
        {
            inputs.Add(item);
        }

        public string ToJason()
        {
            string json = "";

            json += "[" + System.Environment.NewLine;

            foreach (Input input in inputs)
                json += input;

            json += "]" + System.Environment.NewLine;


            return json;
        }
    }
}

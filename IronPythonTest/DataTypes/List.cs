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

        //public override string ToJason()
        //{
        //    string json = "";

        //    json += "{" + System.Environment.NewLine;

        //    json += "type: " + type + "," + System.Environment.NewLine;
        //    json += "label: " + label + "," + System.Environment.NewLine;
        //    json += "id: " + id + "," + System.Environment.NewLine;

        //    json += "inputs: [" + System.Environment.NewLine;

        //    foreach (Input input in inputs)
        //        json += input.ToJason() + "," + System.Environment.NewLine;

        //    json += "]" + System.Environment.NewLine;


        //    json += "}";

        //    return json;
        //}
    }
}

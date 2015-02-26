using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTypes
{
    public class NonAtomic : Input
    {
        protected List<Input> inputs = new List<Input>();

        public virtual void Add(Input item) { }

    }
}

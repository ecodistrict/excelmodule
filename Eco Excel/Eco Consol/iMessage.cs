using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Eco_Consol
{
    interface iMessage
    {
        GlobalParameters.EMethod Method { get; set; }
        GlobalParameters.EType Type { get; set; }

        string ToJsonMessage();


    }
}

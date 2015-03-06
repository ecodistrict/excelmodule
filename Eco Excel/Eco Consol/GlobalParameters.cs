using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Eco_Consol
{
    public static class GlobalParameters
    {
        public enum EMethod
        {
            GetModules,
            SelectModel,
            StartModel
        }

        public enum EType
        {
            Request,
            Response,
            Result
        }
    }
}

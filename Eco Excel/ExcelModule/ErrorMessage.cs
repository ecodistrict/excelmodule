﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ecodistrict.Excel
{
    public class ErrorMessageEventArg:EventArgs
    {
        public string SourceFunction;
        public string Message;
        public Exception Exception=null;
    }
}
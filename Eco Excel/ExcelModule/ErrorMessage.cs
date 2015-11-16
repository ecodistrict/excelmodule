using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ecodistrict.Excel
{
    /// <summary>
    /// Error message event argument that includes name of the function 
    /// that caused the error, the error message and optionally the exception object itself</summary>
    public class ErrorMessageEventArg:EventArgs
    {
        /// <summary>
        /// Name of the function that caused the error
        /// </summary>
        public string SourceFunction;
        /// <summary>
        /// Error message
        /// </summary>
        public string Message;
        //The exception object (optionally included, test for null)
        public Exception Exception=null;
    }
}

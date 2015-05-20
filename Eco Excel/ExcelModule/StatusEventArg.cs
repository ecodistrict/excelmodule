using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ecodistrict.Excel
{
    /// <summary>
    /// Used for sending status messages. Inherites from EventArgs
    /// </summary>
    public class StatusEventArg:EventArgs
    {
        /// <summary>
        /// string Statusmessage
        /// </summary>
        public string StatusMessage;
    }
}

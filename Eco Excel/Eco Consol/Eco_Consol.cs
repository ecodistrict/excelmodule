using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eco_Consol
{
    /// <summary> 
    /// General consol applikation that handles comunication with the dashboard and the Excel class.
    /// Using the configuration file the application connects to the IMB Hub and subscribes to the dashbord
    /// events. After that the application waits for the dashboard to send one of the following commands 
    /// describing what information it wants: 
    /// - Getmodels, which is a request for what Kpis are available from the application.
    /// - SelectModel, which is a request for what variables is needed to start a certain Kpi 
    /// - StartModel, a request to calculate as certain Kpi using the variables passed whith  the request.
    /// The application does the following in turn:
    /// 1. Reads the configuration file
    /// </summary> 
    [System.Runtime.CompilerServices.CompilerGenerated]
    class NamespaceDoc
    {

    }  
}

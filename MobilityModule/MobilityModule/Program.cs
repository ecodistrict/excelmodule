using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobilityModule
{
    class Program
    {
        static void Main(string[] args)
        {
            MobilityModule module = new MobilityModule();

            try
            {
                if (!module.Init("IMB_config.yaml", "Module_config.yaml"))
                {
                    Console.WriteLine("Could not read module settings");
                    return;
                }
                
                if (module.ConnectToServer())
                {
                    Console.WriteLine(">> Press return to close connection");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("**** Errors detected! ****");
                    Console.WriteLine(">> Press return to close");
                    Console.ReadLine();
                    module.Close();
                }
            }
            finally
            {
                module.Close();
                module = null;
            }
        }
    }
}

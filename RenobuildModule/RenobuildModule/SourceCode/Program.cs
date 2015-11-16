using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RenobuildModule
{
    class Program
    {


        static void Main(string[] args)
        {
            RenobuildModule module = new RenobuildModule();

            try
            {
                bool startupStatus = true;

                if (!module.Init("IMB_config.yaml", "Module_config.yaml"))
                {
                    Console.WriteLine("Could not read module settings");
                    startupStatus = false;
                }
                
                startupStatus = module.ConnectToServer();

                if (startupStatus)
                {
                    Console.WriteLine(">> Press return to close connection");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("**** Errors detected! ****");
                    Console.WriteLine(">> Press return to close");
                    Console.ReadLine();
                    if (module != null)
                        module.Close();
                }
            }
            finally
            {
                if (module != null)
                    module.Close();
            }

            return;
        }

    }
}

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
                    startupStatus = false;
                }

                if (!module.ConnectToServer())
                {
                    Console.WriteLine("Could not connect to the IMB-hub");
                    startupStatus = false;
                }
                else
                {
                    Console.WriteLine("Connected..");
                }

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
                    module.Close();
                }
            }
            finally
            {
                module.Close();
            }
        }
    }
}

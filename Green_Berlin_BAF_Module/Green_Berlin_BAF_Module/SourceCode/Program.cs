using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Green_BerlinBAF_Module
{
    class Program
    {
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
            Green_BerlinBAF_Module module = new Green_BerlinBAF_Module();

            try
            {
                bool startupStatus = true;

                if (!module.Init("Config/IMB_config.yaml", "Config/Module_config.yaml"))
                {
                    Console.WriteLine("Could not read module settings");
                    startupStatus = false;
                }

                if (!module.ConnectToServer())
                {
                    Console.WriteLine("Could not connect to the IMB-hub");
                    startupStatus = false;
                }
                else
                {
                    Console.WriteLine("Connected to IMB-hub..");
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

        private static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            Console.WriteLine(args.ToString());
            return null;
        }
    }
}

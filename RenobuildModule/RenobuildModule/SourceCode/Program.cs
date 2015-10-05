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

                    //bool quit = false;
                    //bool testOk = false;
                    //do
                    //{
                    //    try
                    //    {
                    //        Ecodistrict.Excel.Reader.ReadLine(5*6000);
                    //        return;
                    //    }
                    //    catch 
                    //    {
                    //        testOk = !module.TestConnection2();
                    //    }
                    //}
                    //while (module.Connected & !quit);
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

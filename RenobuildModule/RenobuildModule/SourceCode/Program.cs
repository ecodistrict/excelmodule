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
                else
                {
                    //Console.WriteLine("connecting {0} Port {1} User:{2} UserId:{3} Federation: {4}",
                    //    ServerAdress, Port, UserName, UserId, Federation);
                }

                //if (!module.Connect())
                //{
                //    Console.WriteLine("Ending program");
                //    startupStatus = false;
                //}
                //else
                //{
                //    Console.WriteLine("Connected..");
                //}

                if (startupStatus)
                {
                    Console.WriteLine(">> Press return to close connection");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("**** Errors detected! ****");
                    Console.WriteLine(">> Press return to close connection");
                    Console.ReadLine();
                    //if (Connection.Connected)
                    //    Connection.Close();
                }

            }
            finally
            {
                //Connection.Close();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cheese_Module
{
    class Program
    {
        static void Main(string[] args)
        {
            var cheeseModule = new CheeseModule();
            try
            {
                bool startupStatus = true;

                cheeseModule.Init();

                if (!cheeseModule.ConnectToServer())
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
                    cheeseModule.Close();
                }
            }
            finally
            {
                cheeseModule.Close();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ecodistrict.Excel
{
    public class Reader
    {
        private static System.Threading.Thread inputThread;
        private static System.Threading.AutoResetEvent getInput, gotInput;
        private static string input;

        static Reader()
        {
            getInput = new System.Threading.AutoResetEvent(false);
            gotInput = new System.Threading.AutoResetEvent(false);
            inputThread = new System.Threading.Thread(reader);
            inputThread.IsBackground = true;
            inputThread.Start();
        }

        private static void reader()
        {
            while (true)
            {
                getInput.WaitOne();
                input = Console.ReadLine();
                gotInput.Set();
            }
        }

        public static string ReadLine(int timeOutMillisecs)
        {
            getInput.Set();
            bool success = gotInput.WaitOne(timeOutMillisecs);
            if (success)
                return input;
            else
                throw new TimeoutException("User did not provide input within the timelimit.");
        }
    }
}

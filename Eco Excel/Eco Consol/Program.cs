using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using EcoExcel;
using IMB;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;


namespace Eco_Consol
{
    class Program
    {
        //Debug
        private const string UrlPath = @"C:\Users\perbe\Documents\EcoDistr\Kod\Eco Excel\ServerInfo.txt";
        private const string XlsPath = @"C:\Users\perbe\Documents\EcoDistr\Kod\Eco Excel\EcoDistrictTest.xlsx";
        //End debug
        private static TConnection Connection { get; set; }
        private static TEventEntry SubscribedEvent { get; set; }
        private static TEventEntry PublichedEvent { get; set; }
        private static string Id { get; set; } //Self.ID
        private static CServerData ServerData { get; set; }
        
        
        static void Main(string[] args)
        {
            //Debug
            JSONtest();
            //End debug

            try
            {
                bool startupStatus = true;
                if (!ReadConfigFile())
                {
                    Console.WriteLine("Ending program");
                    startupStatus = false;
                }

                if (!ConnectToServer())
                {
                    Console.WriteLine("Ending program");
                    startupStatus = false;
                }
                if (!startupStatus)
                {
                    Console.WriteLine(">> Press return to close connection");
                    Console.ReadLine();
                }

        }
            finally
            {
                Connection.Close();
            }
        }

        private static bool ConnectToServer()
        {
            bool res=true;
            Connection = new TConnection("localhost", 4000, "TNODemo", 0, TConnection.DefaultFederation);
            try
            {
                if (Connection.Connected)
                {
                    SubscribedEvent = Connection.Subscribe("SubscribedEventName");
                    PublichedEvent = Connection.Publish("PublishedEventName");

                    // set event handler for change object on subscribedEvent
                    SubscribedEvent.OnNormalEvent += SubscribedEvent_OnNormalEvent;
                }
                else
                {
                    Console.WriteLine("## NOT connected");
                    res = false;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                res = false;
            }
            return res;
        }

        private static GetModulesMessageRequest JSONtest()
        {
            const string cStartModel = "startModel";
            const string cSelectModel = "selectModel";
            const string cGetModels = "getModels";

            var message = File.ReadAllText(@"L:\EcoDistr\Kod\Eco Excel\" + "getModels.txt");
            //var message = File.ReadAllText(@"L:\EcoDistr\Kod\Eco Excel\" + "selectModel.txt");
            //var message = File.ReadAllText(@"L:\EcoDistr\Kod\Eco Excel\" + "startModel.txt");
            
            dynamic array = JsonConvert.DeserializeObject(message);

            string type, method;
            var c = new List<object>();

            type = array.type;
            
            if(type!="request")
                return null;
            
            method = array.method;
            var reqMsg = new GetModulesMessageRequest();
            

            switch (method)
            {
                case cGetModels:
                {
                    //reqMsg.Method = GetModulesMessageRequest.eMethod.GetModules;
                    break;
                }
                case cSelectModel:
                {
                    //reqMsg.Method=CRequestMessage.eMethod.SelectModel;
                    break;
                }
                case cStartModel:
                {
                    //reqMsg.Method = CRequestMessage.eMethod.StartModel;
                    break;
                }
            }

            if(array.parameters!=null)
                foreach (var item in array.parameters)
                {
                    foreach (var str in item)
                    {
                        c.Add(str);
                    }
                }

            return reqMsg;

        }


         static void SubscribedEvent_OnNormalEvent(TEventEntry aEvent, IMB.ByteBuffers.TByteBuffer aPayload)
        {
            const string cType = "type";
            const string cMethod = "method";
            const string cParameters = "parameters";

            string message= System.Text.Encoding.UTF8.GetString(aPayload.Buffer);

           

            var reader = new JsonTextReader(new StringReader(message));
            while (reader.Read())
            {
                if (reader.Value != null)
                    Console.WriteLine("Token: {0}, Value: {1}", reader.TokenType, reader.Value);
                else
                    Console.WriteLine("Token: {0}", reader.TokenType);
            }
            string method = "";
             var moduleID = "12";

            switch (method)
            {
                case "getModules":
                    SendKPIList();
                    break;
                case "selectModules":
                    if (moduleID == Id)
                    {
                        SendSelectInfo();
                    }
                    break;
                case "startModule":
                    if (moduleID == Id)
                    {
                        SendStatus("Processing");
                        StartModule();
                        SendResult();
                        SendStatus("Success");
                    }
                    break;
            }
        }

        private static void SendResult()
        {
            throw new NotImplementedException();
        }

        private static void SendStatus(string processing)
        {
            throw new NotImplementedException();
        }

        private static void SendSelectInfo()
        {
            throw new NotImplementedException();
        }

        private static void SendKPIList()
        {
            var sb = new StringBuilder();
            var sw = new StringWriter(sb);
            
            using(var jtw=new JsonTextWriter(sw))
            {
                jtw.WriteStartObject();
                jtw.WritePropertyName("Kpi");
                jtw.WriteValue("energy-kpi");
                jtw.WritePropertyName("Kpi");
                jtw.WriteValue("ghg-kpi");
                jtw.WriteEnd();
                jtw.WriteEndObject();
            }
            PublichedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, GetBytes(sb.ToString()));
        }

        private static void StartModule()
        {
            RunExcel();
            //SendResult();
        }

        private static void RunExcel()
        {
            var excel = new CExcel(XlsPath);
        }

        private static bool ReadConfigFile()
        {
            try
            {
                ServerData=new CServerData();
                using (var fs = new FileStream(UrlPath, FileMode.Open, FileAccess.Read))
                {
                    using (var sr = new StreamReader(fs))
                    {
                        ServerData.ServerAdress = sr.ReadLine();
                        ServerData.Port = int.Parse(sr.ReadLine());
                        ServerData.UserId = sr.ReadLine();
                        ServerData.UserName = sr.ReadLine();
                        ServerData.Federation = sr.ReadLine();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        public static byte[] GetBytes(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length); //JSon??
            return bytes;
        }

        private static string[] JSonConv(string jsonStr)
        {
            int count = 0;
            int count2 = 0;
            int inOrOut = 0;
            int nRecords = 1;
            var reader = new JsonTextReader(new StringReader(jsonStr));
            string[] rawData = new string[5];
            while (reader.Read())
            {
                if (reader.Value != null)
                    if (inOrOut == 1)
                    {
                        if (count == 6)
                        {
                            nRecords++;
                            Array.Resize(ref rawData, nRecords);
                            //textBox1.Text += "\r\n";
                            count = 0;
                        }
                        rawData[count2] += reader.Value + ","; //+"\r\n"
                        inOrOut = 0;
                        count++;
                        if (count2 == 500)
                        {
                            //MessageBox.Show(rawData[499]);
                        }
                    }
                    else
                    {
                        inOrOut = 1;
                    }
            }
            return rawData;
        }
    }
}

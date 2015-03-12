using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using EcoExcel;
using IMB;
using Microsoft.Office.Interop.Excel;
using IronPython;
using Ecodistrict.Messaging;
using IronPython.Hosting;


namespace Eco_Consol
{
    class Program
    {
        //Debug
        private const string UrlPath = @"C:\Users\perbe\Documents\EcoDistr\Kod\Eco Excel\ServerInfo.txt";
        //End debug
        private static TConnection Connection { get; set; }
        private static TEventEntry SubscribedEvent { get; set; }
        private static TEventEntry PublichedEvent { get; set; }
        private static string Id { get; set; } //Self.ID
        private static ServerInfo ServerData { get; set; }
        private static dynamic Config { get; set; }

        
        static void Main(string[] args)
        {    
            ReadConfiguration();

#if (DEBUG)
            DebugReadTestFiles();
#endif
            try
            {
                bool startupStatus = true;
                if (!ReadServerConfigFile())
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

        private static void DebugReadTestFiles()
        {
           //Receive and Respond to getModels
            string msg = File.ReadAllText(@"..\..\TestFiles\getModels.txt", new UTF8Encoding());
            GetModelsRequest gmReq= ConvertEcoString(msg) as GetModelsRequest;
            var kpiList = new List<string>();
            foreach (var item in Config.kpiList)
                kpiList.Add(item);
            GetModelsResponse gmRes=new GetModelsResponse(Config.name,Config.moduleId,Config.description,kpiList);
            var gmResString=Ecodistrict.Messaging.Serialize.Message(gmRes);
           //Send gmResString

            
            //Receive and response to selectModel
            msg = File.ReadAllText(@"..\..\TestFiles\selectModel.txt", new UTF8Encoding());
            SelectModelRequest smReq = ConvertEcoString(msg) as SelectModelRequest;
            string variantId = smReq.variantId;
            string kpiId = smReq.kpiId;
            SelectModelResponse smRes = new SelectModelResponse(Config.moduleId, variantId,kpiId,Config.input_specification());
            var smResString = Ecodistrict.Messaging.Serialize.Message(smRes);
            //Send smResString

            //Receive and response to startModule
            msg = File.ReadAllText(@"..\..\TestFiles\startModel.txt", new UTF8Encoding());
            StartModelRequest stmReq = ConvertEcoString(msg) as StartModelRequest;
            string smVariantId = stmReq.variantId;
            string smkpiId = smReq.kpiId;
            StartModelResponse stmResp=new StartModelResponse(Config.moduleId,smVariantId,smkpiId,ModelStatus.Processing);
            var stmRespString=Ecodistrict.Messaging.Serialize.Message(stmResp);
            //Send stmRespString

            CExcel exls;
            Outputs _outputs=null;
            try
            {


                if (File.Exists(Config.path))
                {
                    exls = new CExcel(Config.path);
                    _outputs = Config.run(stmReq.inputData, exls);
                }
                else
                {
                    Console.WriteLine("Excelfile <{0}> not found", Config.path);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                exls = null;    
            }

            ModelResult _result = new ModelResult(Config.moduleId, smVariantId, smkpiId,_outputs);
            var modResString=Ecodistrict.Messaging.Serialize.Message(_result);
            
            StartModelResponse stmResp2 = new StartModelResponse(Config.moduleId, smVariantId, smkpiId, ModelStatus.Success);
            var stmRespString2 = Ecodistrict.Messaging.Serialize.Message(stmResp);
            //Send stmRespString2

        }
        /// <summary>
        /// 
        /// </summary>
        private static void ReadConfiguration()
        {
            var ipy = Python.CreateRuntime();
            Config = ipy.UseFile("../ModuleConfig.py");

            Ecodistrict.Messaging.InputSpecification inputSpec = Config.input_specification();

            //var myString=Ecodistrict.Messaging.Serialize.InputSpecification(inputSpec,true);
            //Gå igenom och beräkna
            //List<Ecodistrict.Messaging.Output> myRes=Config.run("startModuleRequest", ExcelObj);

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

       


         static void SubscribedEvent_OnNormalEvent(TEventEntry aEvent, IMB.ByteBuffers.TByteBuffer aPayload)
        {
            string message= System.Text.Encoding.UTF8.GetString(aPayload.Buffer);

            ConvertEcoString(message);

        }

         private static IMessage ConvertEcoString(string message)
         {
             IMessage iMessage = Deserialize.JsonMessage(message);
             return iMessage;

             //if (iMessage is Ecodistrict.Messaging.GetModelsRequest)
             //{
             //    GetModelsRequest gmr = iMessage as GetModelsRequest;
             //    var gmresp = new Ecodistrict.Messaging.GetModelsResponse(Config.name, Config.moduleId, Config.description, Config.kpiList);
             //    var txt = Ecodistrict.Messaging.Serialize.Message(gmresp);
             //}
             //else if (iMessage is Ecodistrict.Messaging.SelectModelRequest)
             //{
             //    SelectModelRequest rq = iMessage as SelectModelRequest;
             //    if (rq.moduleId == Id)
             //    {

             //    }


             //}
             //else if (iMessage is Ecodistrict.Messaging.StartModelRequest)
             //{
             //    StartModelRequest smr = iMessage as StartModelRequest;
             //    //Här finnds indatalistan
                 
             //}
             //else
             //{
             //}

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
            
            //using(var jtw=new JsonTextWriter(sw))
            //{
            //    jtw.WriteStartObject();
            //    jtw.WritePropertyName("Kpi");
            //    jtw.WriteValue("energy-kpi");
            //    jtw.WritePropertyName("Kpi");
            //    jtw.WriteValue("ghg-kpi");
            //    jtw.WriteEnd();
            //    jtw.WriteEndObject();
            //}
            PublichedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, GetBytes(sb.ToString()));
        }


        private static bool ReadServerConfigFile()
        {
            try
            {
                ServerData=new ServerInfo();
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
    }
}

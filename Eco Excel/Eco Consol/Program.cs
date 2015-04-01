
//#define NoServer

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;
using System.Text;
using System.Threading;
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
        private static string SubScribedEventName { get; set; }
        private static string PublishedEventName { get; set; }
        private static TConnection Connection { get; set; }
        private static TEventEntry SubscribedEvent { get; set; }
        private static TEventEntry PublichedEvent { get; set; }
        private static ServerInfo serverInfo { get; set; }
        private static dynamic Config { get; set; }

        /// <summary>
        /// Main routine 
        /// </summary>
        /// <param name="args">Not used</param>
        static void Main(string[] args)
        {    
            
           

            try
            {
                bool startupStatus = true;

                if (!ReadConfigurationFile())
                {
                    startupStatus = false;
                }
#if (NoServer)
                DebugReadTestFiles();
#endif
                if (!GetServerConfiguration())
                {
                    Console.WriteLine("Error reading serverinfo from Config file!");
                    startupStatus = false;
                }
                else
                {
                    Console.WriteLine("connecting {0} Port {1} User:{2} UserId:{3} Federation: {4}",
                        serverInfo.ServerAdress, serverInfo.Port, serverInfo.UserName, serverInfo.UserId,
                        serverInfo.Federation);
                }

                if (!ConnectToServer())
                {
                    Console.WriteLine("Ending program");
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
            GetModulesRequest gmReq = Deserialize.JsonString(msg) as GetModulesRequest;
            var kpiList = new List<string>();
            foreach (var item in Config.kpiList)
                kpiList.Add(item);
            GetModulesResponse gmRes=new GetModulesResponse(Config.name,Config.moduleId,Config.description,kpiList);
            var gmResString=Ecodistrict.Messaging.Serialize.ToJsonString(gmRes, true);
           //Send gmResString

            
            //Receive and response to selectModel
            msg = File.ReadAllText(@"..\..\TestFiles\selectModel.txt", new UTF8Encoding());
            SelectModuleRequest smReq = Deserialize.JsonString(msg) as SelectModuleRequest;
            string variantId = smReq.variantId;
            string kpiId = smReq.kpiId;
            if (Config.kpiId_Exists(kpiId))
            {
                SelectModuleResponse smRes = new SelectModuleResponse(Config.moduleId, variantId,kpiId,Config.input_specification(kpiId));
                var smResString = Ecodistrict.Messaging.Serialize.ToJsonString(smRes, true);
            }
            //Send smResString

            //Receive and response to startModule
            msg = File.ReadAllText(@"..\..\TestFiles\startModel.txt", new UTF8Encoding());
            StartModuleRequest stmReq = Deserialize.JsonString(msg) as StartModuleRequest;
            string smVariantId = stmReq.variantId;
            string smkpiId = smReq.kpiId;
            StartModuleResponse stmResp=new StartModuleResponse(Config.moduleId,smVariantId,smkpiId,ModuleStatus.Processing);
            var stmRespString=Ecodistrict.Messaging.Serialize.ToJsonString(stmResp, true);
            //Send stmRespString

            CExcel exls;
            Outputs _outputs=null;
            try
            {


                if (File.Exists(Config.path))
                {
                    exls = new CExcel(Config.path);
                    _outputs = Config.run(stmReq.inputData,smkpiId, exls);
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

            ModuleResult _result = new ModuleResult(Config.moduleId, smVariantId, smkpiId,_outputs);
            var modResString=Ecodistrict.Messaging.Serialize.ToJsonString(_result, true);
            
            StartModuleResponse stmResp2 = new StartModuleResponse(Config.moduleId, smVariantId, smkpiId, ModuleStatus.Success);
            var stmRespString2 = Ecodistrict.Messaging.Serialize.ToJsonString(stmResp,true);
            //Send stmRespString2

        }
       
        private static bool ReadConfigurationFile()
        {
            const string fileName = "ModuleConfig.py";
            
            try
            {
                var exeName = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var exeDirectory = Path.GetDirectoryName(exeName);
                var filePath = Path.Combine(exeDirectory, fileName);
                //var filePathPy = "../" + fileName;

                var fi = new FileInfo(filePath);
                if (!fi.Exists)
                {
                    Console.WriteLine("Can´t find file: {0}",fi.FullName);
                    return false;
                }
                var ipy = Python.CreateRuntime();
                Config = ipy.UseFile(filePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        private static bool GetServerConfiguration()
        {
            try
            {
                SubScribedEventName = Config.subScribedEvent;
                PublishedEventName = Config.publishedEvent;
                serverInfo=new ServerInfo();
                serverInfo.ServerAdress = Config.serverAdress;
                serverInfo.Port = (int) Config.port;
                serverInfo.UserId = Config.userId;
                serverInfo.UserName = Config.userName;
                serverInfo.Federation = Config.federation;

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        private static bool ConnectToServer()
        {
            bool res=true;
            Connection = new TConnection(serverInfo.ServerAdress, serverInfo.Port,serverInfo.UserName , serverInfo.UserId,serverInfo.Federation);
            try
            {
                if (Connection.Connected)
                {
                    SubscribedEvent = Connection.Subscribe(SubScribedEventName);
                    PublichedEvent = Connection.Publish(PublishedEventName);

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
             var nByteArr = new byte[aPayload.Buffer.Length - 12];
             Buffer.BlockCopy(aPayload.Buffer, 12, nByteArr, 0, aPayload.Buffer.Length - 12);

             IMessage iMessage = Ecodistrict.Messaging.Deserialize.JsonByteArr(nByteArr);

             if (iMessage is Ecodistrict.Messaging.GetModulesRequest)
             {
                 SendGetModulesResponse();
             }
             else if (iMessage is SelectModuleRequest)
             {
                 var SMR = iMessage as SelectModuleRequest;
                 if(Config.moduleId==SMR.moduleId)
                 {
                     SendSelectModuleResponse(SMR);    
                 }
             }
             else if (iMessage is StartModuleRequest)
             {
                 var SMR = iMessage as StartModuleRequest;
                 if (Config.moduleId == SMR.moduleId)
                 {
                    
                     SendModuleResult(SMR);
                 }
             }
        }

        private static bool SendGetModulesResponse()
        {
            try
            {
                var kpiList = new List<string>();
                foreach (var item in Config.kpiList)
                    kpiList.Add(item);
                GetModulesResponse gmRes = new GetModulesResponse(Config.name, Config.moduleId, Config.description, kpiList);

                var msgBytes = Serialize.ToJsonByteArr(gmRes);
                PublichedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, msgBytes);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool SendSelectModuleResponse(SelectModuleRequest request)
        {
            var variantId = request.variantId;
            var kpiId = request.kpiId;
            SelectModuleResponse smResponse=new SelectModuleResponse(Config.moduleId,variantId,kpiId,Config.input_specification(kpiId));
            var smRespBytes = Serialize.ToJsonByteArr(smResponse);
            PublichedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, smRespBytes);
            return true;
        }

        private static bool SendModuleResult(StartModuleRequest request)
        {
            var smr =new StartModuleResponse(Config.moduleId,request.variantId,request.kpiId,ModuleStatus.Processing);
            var respBytes = Serialize.ToJsonByteArr(smr);
            PublichedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, respBytes);

            CExcel exls=null;
            Outputs outputs = null;


            try
            {
                if (File.Exists(Config.path))
                {
                    exls = new CExcel(Config.path);
                    outputs = Config.run(request.inputData,request.kpiId ,exls);
                }
                else
                {
                    Console.WriteLine("Excelfile <{0}> not found", Config.path);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
                if(exls!=null)
                    exls.CloseExcel();
                exls = null;
            }

            ModuleResult result = new ModuleResult(Config.moduleId, request.variantId, request.kpiId, outputs);
            var modResString=Ecodistrict.Messaging.Serialize.ToJsonString(result, true);
            var modResBytes = Ecodistrict.Messaging.Serialize.ToJsonByteArr(result);
            var mres = (ModuleResult) Deserialize.JsonByteArr(modResBytes);
            var mres2 = (ModuleResult)Deserialize.JsonString(modResString);

            PublichedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, modResBytes);

            StartModuleResponse stmResp2 = new StartModuleResponse(Config.moduleId, request.variantId, request.kpiId, ModuleStatus.Success);
            var stmRespString2 = Ecodistrict.Messaging.Serialize.ToJsonByteArr(stmResp2);
            PublichedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, stmRespString2);
            return true;
        }


    }
}


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
using IMB.ByteBuffers;
using Microsoft.Office.Interop.Excel;
using IronPython;
using Ecodistrict.Messaging;
using IronPython.Hosting;


namespace Eco_Consol
{
    /// <summary>
    /// Eco Excel application.<br/>
    /// Module that connects an Excel document to the ECOdistr-ICT dashboard.
    /// <remarks>
    /// The application is a console application that starts with reading a
    /// configuration file, ModuleConfig.py, that should reside in the same directory as the application.
    /// Failure reading the file will result in that the application closes down after having written
    /// a message describing the error to the consol.<br/>
    /// After having read the congiguration file the application tries to connect to the IMB hub using the 
    /// parameters: serverAdress, port, userId, userName and federation.<br/>
    /// If it succeeds it uses the information from the config file, subscribedEvent, to subscribe for dashboard
    /// calls.<br/>
    /// 
    /// 
    /// </remarks>
    /// </summary>
    public class Program
    {
        /// <summary>
        /// The name of the IMB subscription the application uses. Read from the configuration file
        /// </summary>
        private static string SubScribedEventName { get; set; }
        /// <summary>
        /// The name the application uses when sending back information to the dashboard
        /// </summary>
        private static string PublishedEventName { get; set; }

        private static TConnection Connection { get; set; }
        private static TEventEntry SubscribedEvent { get; set; }
        private static TEventEntry PublishedEvent { get; set; }
        private static dynamic Config { get; set; }

        private static string ServerAdress { get; set; }
        private static int Port { get; set; }
        private static int UserId { get; set; }
        private static string ModuleId { get; set; }
        private static string UserName { get; set; }
        private static string Federation { get; set; }

        /// <summary>
        /// Main routine 
        /// </summary>
        /// <remarks>
        /// 1. Reads the config file<br/>
        /// 2. Connects to IMB hub<br/>
        /// 3. Waits and answers calls from dashboard until Return is peressed on the keyboard<br/>
        /// 
        /// If anything fails during startup the application sends an errormessage to the console and finishes
        /// 
        /// </remarks>
        /// <param name="args">Not used</param>
        static void Main(string[] args)
        {    
            try
            {
                bool startupStatus = true;

                if (!ReadConfigurationFile())
                {
                    Console.WriteLine("Error reading serverinfo from Config file!");
                    startupStatus = false;
                }
                else
                {
                    Console.WriteLine("connecting {0} Port {1} User:{2} UserId:{3} Federation: {4}",
                        ServerAdress, Port, UserName, UserId, Federation);
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
                    if (Connection.Connected) 
                        Connection.Close();
                }

        }
            finally
            {
                Connection.Close();
            }
        }
       
        /// <summary>
        /// Reads the configuration file
        /// </summary>
        /// <returns></returns>
        private static bool ReadConfigurationFile()
        {
            const string serverFileName="ServerConfig.py";
            const string moduleFileName = "ModuleConfig.py";
            
            try
            {
                var filePathServerConfig = Path.Combine(GetExeDirectory(), serverFileName);
                var filePathModuleConfig = Path.Combine(GetExeDirectory(), moduleFileName);


                var fi = new FileInfo(filePathServerConfig);
                if (!fi.Exists)
                {

                    Console.WriteLine("Can´t find file: {0}", fi.FullName);
                    return false;
                }
                else
                {
                    try
                    {
                        var ipyS = Python.CreateRuntime();
                        dynamic ServerConfig = ipyS.UseFile(filePathServerConfig);
                        ServerAdress = ServerConfig.serverAdress;
                        Port = ServerConfig.port;
                        UserId = ServerConfig.userId;
                        UserName = ServerConfig.userName;
                        Federation = ServerConfig.federation;
                        SubScribedEventName = ServerConfig.subScribedEvent;
                        PublishedEventName = ServerConfig.publishedEvent;
                        ServerConfig = null;
                        ipyS = null;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        return false;
                    }
                }
                

                
                fi = new FileInfo(filePathModuleConfig);
                if (!fi.Exists)
                {
                    Console.WriteLine("Can´t find file: {0}",fi.FullName);
                    return false;
                }
                var ipy = Python.CreateRuntime();
                Config = ipy.UseFile(filePathModuleConfig);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Reads the serverconfiguration settings from configuration file 
        /// </summary>
        /// <returns></returns>
        //private static bool GetServerConfiguration()
        //{
        //    try
        //    {
        //        SubScribedEventName = Config.subScribedEvent;
        //        PublishedEventName = Config.publishedEvent;
        //        ServerAdress = Config.serverAdress;
        //        Port = (int) Config.port;
        //        UserId = Config.userId;
        //        UserName = Config.userName;
        //        ModuleId = Config.moduleId;
        //        Federation = Config.federation;

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //        return false;
        //    }
        //}

        /// <summary>
        /// Connects to the IMB server using the settings found in the config file.
        /// </summary>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private static bool ConnectToServer()
        {
            bool res=true;
            Connection = new TConnection(ServerAdress, Port, UserName , UserId, Federation);
            try
            {
                if (Connection.Connected)
                {
                    SubscribedEvent = Connection.Subscribe(SubScribedEventName);
                    PublishedEvent = Connection.Publish(PublishedEventName);

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

         /// <summary>
         /// Calls from dashboard handles here
         /// </summary>
         /// <param name="aEvent">EventInformation, TEventEntry</param>
         /// <param name="aPayload">Payload, TByteBuffer</param>
         static void SubscribedEvent_OnNormalEvent(TEventEntry aEvent, IMB.ByteBuffers.TByteBuffer aPayload)
         {

             String msg = "";
             aPayload.Read(out msg);

             //var nByteArr = new byte[aPayload.Buffer.Length - 12];
             //Buffer.BlockCopy(aPayload.Buffer, 12, nByteArr, 0, aPayload.Buffer.Length - 12);

             IMessage iMessage = Ecodistrict.Messaging.Deserialize.JsonString(msg);

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

        /// <summary>
        /// Handles GetModules request
        /// </summary>
         /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private static bool SendGetModulesResponse()
        {
            try
            {
                var kpiList = new List<string>();
                foreach (var item in Config.kpiList)
                    kpiList.Add(item);
                GetModulesResponse gmRes = new GetModulesResponse(Config.name, Config.moduleId, Config.description, kpiList);

                var str = Serialize.ToJsonString(gmRes);
                var payload = new TByteBuffer();
                payload.Prepare(str);
                payload.PrepareApply();
                payload.QWrite(str);
                PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);
                
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Handles SelectModule requests
        /// </summary>
        /// <param name="request">Request, SelectModuleRequest</param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private static bool SendSelectModuleResponse(SelectModuleRequest request)
        {
            var variantId = request.variantId;
            var kpiId = request.kpiId;
            var smResponse=new SelectModuleResponse(Config.moduleId,variantId,kpiId,Config.input_specification(kpiId));
            
            var str = Serialize.ToJsonString(smResponse);
            var payload = new TByteBuffer();
            payload.Prepare(str);
            payload.PrepareApply();
            payload.QWrite(str);
            PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);

            return true;
        }

        /// <summary>
        /// Handles StartModule requests
        /// </summary>
        /// <param name="request">Request, StartModuleRequest</param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private static bool SendModuleResult(StartModuleRequest request)
        {
            var smr =new StartModuleResponse(Config.moduleId,request.variantId,request.kpiId,ModuleStatus.Processing);

            var str = Serialize.ToJsonString(smr);
            var payload = new TByteBuffer();
            payload.Prepare(str);
            payload.PrepareApply();
            payload.QWrite(str);
            PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);


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
                {
                    exls.CloseExcel();
                    exls = null;
                }
            }

            ModuleResult result = new ModuleResult(Config.moduleId, request.variantId, request.kpiId, outputs);
            str=Ecodistrict.Messaging.Serialize.ToJsonString(result);
            payload = new TByteBuffer(); 
            payload.Prepare(str);
            payload.PrepareApply();
            payload.QWrite(str);
            PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);

            StartModuleResponse stmResp2 = new StartModuleResponse(Config.moduleId, request.variantId, request.kpiId, ModuleStatus.Success);
            str = Ecodistrict.Messaging.Serialize.ToJsonString(stmResp2);
            payload = new TByteBuffer();
            payload.Prepare(str);
            payload.PrepareApply();
            payload.QWrite(str);
            PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);

            return true;
        }

        private static string GetExeDirectory()
        {
            var exeName = System.Reflection.Assembly.GetExecutingAssembly().Location;
            return Path.GetDirectoryName(exeName);
        }

    }
}

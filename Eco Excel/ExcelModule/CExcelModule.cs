using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Yaml.Serialization;
using IMB;
using Ecodistrict.Messaging;
using Ecodistrict.Messaging.Requests;
using Ecodistrict.Messaging.Responses;
using Ecodistrict.Messaging.Results;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Data.Odbc;
using System.Data;
using System.Security.Cryptography;
using System.Security;
using System.Security.Permissions;

namespace Ecodistrict.Excel
{
    /// <summary>
    /// Event handler used for Error reporting
    /// </summary>
    /// <param name="sender">reference to the object that raised the event</param>
    /// <param name="e">ErrorMessageEventArg that is inherited from EventArgs</param>
    public delegate void ErrorEventHandler(object sender, ErrorMessageEventArg e);
    /// <summary>
    /// 
    /// </summary>
    /// <param name="sender">reference to the object that raised the event</param>
    /// <param name="e">StatusEventArg, inherites from EventArg and includes a statusmessage string</param>
    public delegate void StatusEventHandler(object sender, StatusEventArg e);

    /// <summary>
    /// The main class that connects to the hub/dashboard and handles the connection with Excel.
    /// This class is abstract and should be be inherited from.
    /// </summary>
    public abstract class CExcelModule
    {
        List<ModuleProcess> _processes;
        List<ModuleProcess> Processes
        {
            get
            {
                if (_processes == null)
                    _processes = new List<ModuleProcess>();

                return _processes;
            }
            set
            {
                _processes = value;
            }
        }
        protected bool useBothVariantAndAsISForVariant = false;
        double timeLimitProcess = 20000;
        System.Timers.Timer checkProcessesTimer = new System.Timers.Timer(20000);
        protected bool useDummyDB = true;

        private void OnCheckProcessesEvent(Object source, System.Timers.ElapsedEventArgs e)
        {
            lock (Processes)
            {
                List<ModuleProcess> ProcessesToRemove = new List<ModuleProcess>();

                foreach (ModuleProcess process in Processes)
                {
                    lock (process)
                    {
                        if (process.TimerExpired)
                            ProcessesToRemove.Add(process);
                    }
                }

                foreach (ModuleProcess process in ProcessesToRemove)
                {
                    lock (process)
                    {
                        Processes.Remove(process);
                        SendStartModuleResponse(process.Request, ModuleStatus.Failed, "Unable to get data from data module");
                        SendErrorMessage(String.Format("Unable to get data from data module for kpi {0}", process.KpiId), "OnCheckProcessesEvent");
                    }
                }
            }

        }

        #region Initialization
        /// <summary>
        /// Creates a new CExcel instance that in turn creates a new instance of Excel.Application
        /// </summary>
        protected CExcelModule()
        {
            try
            {
                ShowOnlyOwnStatus = true;
                ExcelApplikation = new CExcel();

            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "CExcel Constructor", exception: ex);
            }
        }
        /// <summary>
        /// If there is an instance of CExcel object it closes down both any opened Excel documents
        ///  (without saving any data) and closes down Excel.Application before closing down the CExcelModule object.
        /// </summary>
        ~CExcelModule()
        {
            Close();
        }
        #endregion

        #region Module Properties
        /// <summary>
        /// UserId used for identification against the server. 
        /// </summary>
        protected int UserId { get; set; }

        /// <summary>
        /// Name of this module.
        /// </summary>
        protected string ModuleName { get; set; }

        /// <summary>
        /// Used to uniquely identify this module 
        /// </summary>
        protected string ModuleId { get; set; }

        /// <summary>
        /// Name of the module owner/responsible
        /// </summary>
        protected string UserName { get; set; }

        /// <summary>
        /// A list of strings with Kpis that the ExcelFile can calculate.
        /// </summary>
        protected List<string> KpiList { get; set; }
        /// <summary>
        /// Description of the module.
        /// </summary>
        protected string Description { get; set; }
        #endregion

        #region Module Init
        protected virtual void Init_IMB(string IMB_config_path)
        {
            try
            {
                var serializer = new YamlSerializer();
                var imb_settings = serializer.DeserializeFromFile(IMB_config_path, typeof(IMB_Settings))[0];

                this.RemoteHost = ((IMB_Settings)imb_settings).remoteHost;
                this.RemotePort = ((IMB_Settings)imb_settings).remotePort;
                this.SubScribedEventName = ((IMB_Settings)imb_settings).subScribedEventName;
                this.PublishedEventName = ((IMB_Settings)imb_settings).publishedEventName;
                this.PublishedDataModuleEventName = ((IMB_Settings)imb_settings).publishedDataModuleEventName;
                this.aCertFile = ((IMB_Settings)imb_settings).aCertFile;
                this.aCertFilePassword = ((IMB_Settings)imb_settings).aCertFilePassword;
                this.aRootCertFile = ((IMB_Settings)imb_settings).aRootCertFile;
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "Init_IMB", exception: ex);
            }
        }

        protected virtual void Init_Module(string Module_config_path)
        {
            try
            {
                var serializer = new YamlSerializer();
                var module_settings = serializer.DeserializeFromFile(Module_config_path, typeof(Module_Settings))[0];

                this.ModuleName = ((Module_Settings)module_settings).name;
                this.Description = ((Module_Settings)module_settings).description;
                this.ModuleId = ((Module_Settings)module_settings).moduleId;
                this.WorkBookPath = ((Module_Settings)module_settings).path;
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "Init_Module", exception: ex);
            }
        }

        public bool Init(string IMB_config_path, string Module_config_path)
        {
            try
            {
                Init_IMB(IMB_config_path);
                Init_Module(Module_config_path);

                //HUB info
                this.UserId = 0;
                this.UserName = this.ModuleName;

                return true;
            }
            catch (Exception ex)
            {
                CExcelModule_ErrorRaised(this, ex);
                return false;
            }
        }
        #endregion

        #region IMB Hub
        public TConnection Connection { get; set; }
        public TEventEntry SubscribedEvent { get; set; }
        public TEventEntry PublishedEvent { get; set; }
        public TEventEntry SubscribedDataEvent { get; set; }
        public TEventEntry PublishedDataEvent { get; set; }
        /// <summary>
        /// The name of the IMB subscription the application uses. Read from the configuration file
        /// </summary>
        protected string SubScribedEventName { get; set; }
        /// <summary>
        /// The name the application uses when sending back information to the dashboard
        /// </summary>
        protected string PublishedEventName { get; set; }
        /// <summary>
        /// The name the application uses when sending back information to the datamodule
        /// </summary>
        protected string PublishedDataModuleEventName { get; set; }
        string aCertFile;
        string aCertFilePassword;
        string aRootCertFile;
        string aPrefix = "ecodistrict";
        protected string RemoteHost { get; set; }
        protected int RemotePort { get; set; }

        /// <summary>
        /// Close any opened Excel dokument and Excel.Application as a preparation for closing down.
        /// If anything more has to be done it can be overrided.
        /// </summary>
        public virtual void Close()
        {
            try
            {
                if (ExcelApplikation != null)
                    ExcelApplikation.CloseExcel();

                ExcelApplikation = null;

                if (Connected)
                {
                    Connection.onDisconnect -= Connection_onDisconnect;
                    Connection.setHeartBeat(-1);

                    checkProcessesTimer.Stop();

                    //Close connection
                    Connection.close();
                    // reset event handler for change object on subscribedEvent
                    if (SubscribedEvent != null)
                    {
                        SubscribedEvent.onString -= SubscribedEvent_onString;
                        SubscribedDataEvent.onString -= SubscribedEvent_onString;
                    }
                }
            }
            catch (Exception ex)
            {
                // reset event handler for change object on subscribedEvent
                if (SubscribedEvent != null)
                {
                    SubscribedEvent.onString -= SubscribedEvent_onString;
                    SubscribedDataEvent.onString -= SubscribedEvent_onString;
                }
                SendErrorMessage(message: ex.Message, sourceFunction: "Close", exception: ex);
            }

            Connection = null;
            SubscribedEvent = null;
            PublishedEvent = null;

        }

        /// <summary>
        /// Connects to hub/Server, prepares the publish event and starts subscription to dashboard events
        /// </summary>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        public virtual bool ConnectToServer()
        {
            bool res = false;

            try
            {
                if (!Connected)
                {
                    SendStatusMessage("Connecting to IMB-hub..");
                    Connection = new TTLSConnection(aCertFile, aCertFilePassword, aRootCertFile,
                        false,
                        ModuleName, UserId,
                        aPrefix,
                        RemoteHost,
                        RemotePort);

                    if (Connection.connected)
                    {
                        SubscribedEvent = Connection.subscribe(SubScribedEventName);
                        SubscribedDataEvent = Connection.subscribe(Connection.privateEventName, false);
                        PublishedEvent = Connection.publish(PublishedEventName);
                        PublishedDataEvent = Connection.publish(PublishedDataModuleEventName);
                        Connection.onDisconnect += Connection_onDisconnect;
                        Connection.setHeartBeat(60000);

                        checkProcessesTimer.Elapsed += OnCheckProcessesEvent;
                        checkProcessesTimer.AutoReset = true;
                        checkProcessesTimer.Enabled = true;
                        checkProcessesTimer.Start();


                        // set event handler for change object on subscribedEvent
                        SubscribedEvent.onString += SubscribedEvent_onString;
                        SubscribedDataEvent.onString += SubscribedEvent_onString;

                        SendStatusMessage("Connected to IMB-hub..");
                        res = true;
                    }
                    else
                    {
                        SendStatusMessage("Could not connect to the IMB-hub..");
                        res = false;
                    }
                }
                else
                {
                    SendStatusMessage("Already connected to the IMB-hub");
                    res = true;
                }

            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "ConnectToServer", exception: ex);
                res = false;
            }

            return res;
        }

        void Connection_onDisconnect(TConnection aConnection)
        {
            SendStatusMessage("IMB connection lost");
            SendStatusMessage("Try to reconnect to IMB..");
            ReConnect(Int32.MaxValue);
        }

        /// <summary>
        /// Try to reconnect to IMB-hub
        /// </summary>
        /// <param name="nrTries">Number of attempts</param>
        /// <param name="msec">Time between tries</param>
        /// <returns></returns>
        bool ReConnect(uint nrTries, int msec = 25000)
        {
            for (uint i = 0; i < nrTries; ++i)
            {
                if (ConnectToServer())
                    return true;

                if (i == nrTries - 1)
                    SendStatusMessage("Could not reconnect to server...");
                else
                {
                    SendStatusMessage(String.Format("Reconnect in {0} seconds", Math.Round(msec / 1000.0, 0)));
                    System.Threading.Thread.Sleep(msec);
                }
            }

            return false;
        }

        public bool Connected
        {
            get
            {
                if (Connection == null)
                    return false;

                return Connection.connected;
            }
        }

        private void Publish(String str)
        {
            if (PublishedEvent != null)
            {
                lock (PublishedEvent)
                {
                    PublishedEvent.signalString(str);
                }
            }
        }

        private void PublishData(String str)
        {
            if (PublishedDataEvent != null)
            {
                lock (PublishedEvent)
                {
                    PublishedDataEvent.signalString(str);
                }
            }
        }
        #endregion

        #region Events
        void SubscribedEvent_onString(TEventEntry aEventEntry, string msg)
        {
            HandleHubMessage(msg);
        }

        protected virtual void CExcelModule_StatusMessage(object sender, StatusEventArg e)
        {
            Console.WriteLine(String.Format("# {0} #\tStatus message:\t{1}", DateTime.Now.ToString(), e.StatusMessage));
        }

        protected virtual void CExcelModule_ErrorRaised(object sender, ErrorMessageEventArg e)
        {
            Console.WriteLine(String.Format("# {0} #\tError message:\t{1}", DateTime.Now.ToString(), e.Message));
            if (e.SourceFunction != null & e.SourceFunction != "")
                Console.WriteLine(String.Format("\tIn source function: {0}", e.SourceFunction));
        }

        protected virtual void CExcelModule_ErrorRaised(object sender, Exception ex)
        {
            ErrorMessageEventArg em = new ErrorMessageEventArg();
            em.Message = ex.Message;
            CExcelModule_ErrorRaised(sender, em);
        }
        #endregion

        #region Excel
        private string _workBookPath;
        /// <summary>
        /// The complete path to the Excedocument document file that the module is going to use (*.xls, *.xlsx)
        /// </summary>
        protected string WorkBookPath
        {
            get
            {
                return _workBookPath;
            }
            set
            {
                _workBookPath = Path.GetFullPath(value);
            }
        }
        /// <summary>
        /// This function is to be inherited. It receives a Kpi name as string and should return the Inputspecification for that Kpi
        /// </summary>
        /// <param name="kpiId">The name of the Kpi</param>
        /// <returns>InputSpecification object that can be serialized and sent to th dashboard</returns>
        public virtual bool OpenWorkbook()
        {
            bool res = true;

            try
            {
                ExcelApplikation.OpenWorkBook(WorkBookPath);

            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "OpenWorkbook", exception: ex);
                res = false;
            }
            return res;

        }
        private CExcel ExcelApplikation { get; set; }
        #endregion

        #region Handle Hub Messages
        protected bool useXLSData = false;
        void HandleHubMessage(string msg)
        {
            var message = Deserialize<IMessage>.JsonString(msg);

            if (message is GetModulesRequest)
                HandleGetModulesRequest(message as GetModulesRequest);
            else if (message is SelectModuleRequest) //Currently not used. The dashboard doesn't send these anymore, since the input specifikation was removed.
                HandleSelectModuleRequest(message as SelectModuleRequest); 
            else if (message is StartModuleRequest)
            {
                //HandleStartModuleRequest2(message as StartModuleRequest);
                //if (useXLSData)
                //    HandleStartModuleRequest2(message as StartModuleRequest);
                //else                //Override only use hard coded
                if (useDummyDB)
                    HandleStartModuleRequestDummy(message as StartModuleRequest);
                else
                    HandleStartModuleRequest(message as StartModuleRequest);
            }
            else if (message is GetDataResponse)
                HandleGetDataResponse(message as GetDataResponse);
            else if (message is GetKpiResultRequest |
                     message is SetKpiResultRequest |
                     message is GetDataRequest)
                ;
            else
                SendStatusMessage(String.Format("Unknown message; method: {0} , type: {1}", message.method, message.type));
        }

        private void HandleGetModulesRequest(GetModulesRequest request)
        {
            if (request != null)
            {
                SendStatusMessage("GetModulesRequest received");
                if (!SendGetModulesResponse())
                    SendErrorMessage(message: "could not send getModulesResponse", sourceFunction: "SubscribedEvent_OnNormalEvent");
            }
        }

        private void HandleSelectModuleRequest(SelectModuleRequest request)
        {
            if (request != null & request.moduleId == ModuleId)
            {
                if (!ShowOnlyOwnStatus)
                    SendStatusMessage("SelectModuleRequest received");

                SendStatusMessage("Handles SelectModuleRequest");
                SendSelectModuleResponse(request);
            }
        }

        private void HandleStartModuleRequestDummy(StartModuleRequest request)
        {
            if (request != null & request.moduleId == ModuleId)
            {

                var realReq = request;

                //Case: Hovsjo id mapping
                if (request.caseId == "56cd5745317e88872c28ec12")
                {
                    switch (request.variantId)
                    {
                        case null: //AsIS
                            //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "asis", userId: "cstb", kpiId: request.kpiId);
                            request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "greenfactoralt1", userId: "cstb", kpiId: "green-area-factor");
                            break;
                        case "56cd61a1317e88872c28ec16": // Alternative 1A
                        case "56d469d64a93972f1f61ad4a": // Alternative 2
                        case "56d469e64a93972f1f61ad4b": // Alternative 3
                        case "56d46a0c4a93972f1f61ad4c": // Alternative 4A
                        case "56d46a354a93972f1f61ad4d": // Alternative 4B
                            request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "greenfactoralt1", userId: "cstb", kpiId: "green-area-factor");
                            break;
                        case "56d6faae4a93972f1f61ae7c": // Alternative Green
                            request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "greenfactoralt2", userId: "cstb", kpiId: "green-area-factor");
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    //request = new StartModuleRequest(this.ModuleId, "test", "truite", "cstb", "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "asis", userId: "cstb", kpiId: "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "greenfactoralt1", userId: "cstb", kpiId: "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "greenfactoralt2", userId: "cstb", kpiId: "green-area-factor");

                    // LCA/LCC
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "lcalccalt1a", userId: "cstb", kpiId: "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "lcalccalt1b", userId: "cstb", kpiId: "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "lcalccalt2", userId: "cstb", kpiId: "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "lcalccalt3", userId: "cstb", kpiId: "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "lcalccalt4a", userId: "cstb", kpiId: "green-area-factor");
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "hovsjo", variantId: "lcalccalt4b", userId: "cstb", kpiId: "green-area-factor");  

                    //Green - Warsaw
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "green_factor", variantId: "asis", userId: "cstb", kpiId: realReq.kpiId);
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "green_factor", variantId: "variant1", userId: "cstb", kpiId: realReq.kpiId);
                    //request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "green_factor", variantId: "variant2", userId: "cstb", kpiId: realReq.kpiId);

                    //Mobility - Warsaw
                    request = new StartModuleRequest(moduleId: this.ModuleId, caseId: "warsaw_mobility", variantId: "variant2", userId: "cstb", kpiId: realReq.kpiId);

                }

                if (!ShowOnlyOwnStatus)
                    SendStatusMessage("StartModuleRequest received");

                SendStatusMessage("Handles StartModuleRequest");

                if (SendStartModuleResponse(realReq, ModuleStatus.Processing, "Accessing data"))
                {
                    var process = new ModuleProcess(realReq, timeLimitProcess);

                    if (useBothVariantAndAsISForVariant & realReq.variantId != null)
                    {
                        //process.As_IS_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "hovsjo", "greenfactoralt1", dataEventId, "cstb");
                        //process.Variant_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "hovsjo", "greenfactoralt2", dataEventId, "cstb");

                        process.As_IS_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "warsaw_mobility", null, Connection.privateEventName, "cstb");
                        //process.Variant_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "warsaw_mobility", "alt1", dataEventId, "cstb");
                        //process.Variant_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "warsaw_mobility", "alt2", dataEventId, "cstb");
                        process.Variant_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "warsaw_mobility", "alt3", Connection.privateEventName, "cstb");


                        //process.As_IS_Request =
                        //    new GetDataRequest(ModuleId, Guid.NewGuid().ToString(), request.caseId, request.variantId, dataEventId);
                        //process.Variant_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), dataEventId);
                    }
                    //else 
                    //    process.Variant_Request = new GetDataRequest(request, Convert.ToString(4), dataEventId);
                    else if (realReq.variantId != null)
                        //else if (request.variantId != "greenfactoralt2")
                        process.Variant_Request = new GetDataRequest(request, Convert.ToString(4), Connection.privateEventName);
                    else
                        process.As_IS_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "warsaw_mobility", null, Connection.privateEventName, "cstb");
                        //process.As_IS_Request = new GetDataRequest(ModuleId, Convert.ToString(4), "hovsjo", "greenfactoralt1", dataEventId, "cstb");
                    //process.As_IS_Request = new GetDataRequest(request, Convert.ToString(4), dataEventId);


                    lock (Processes)
                    {
                        Processes.Add(process);
                    }

                    if (process.As_IS_Request != null)
                    {
                        SendStatusMessage("GetDataRequest sent for AS-IS");
                        SendDataModuleMessage(process.As_IS_Request);
                    }

                    if (process.Variant_Request != null)
                    {
                        SendStatusMessage("GetDataRequest sent for Variant");
                        SendDataModuleMessage(process.Variant_Request);
                    }
                }

                return;

            }

        }

        private void HandleStartModuleRequest(StartModuleRequest request)
        {
            if (request != null & request.moduleId == ModuleId)
            {

                if (!ShowOnlyOwnStatus)
                    SendStatusMessage("StartModuleRequest received");

                SendStatusMessage("Handles StartModuleRequest");

                if (SendStartModuleResponse(request, ModuleStatus.Processing, "Accessing data"))
                {
                    var process = new ModuleProcess(request, timeLimitProcess);

                    if (useBothVariantAndAsISForVariant & request.variantId != null)
                    {
                        process.As_IS_Request =
                            new GetDataRequest(ModuleId, Guid.NewGuid().ToString(), request.caseId, null, Connection.privateEventName, request.userId);
                        process.Variant_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), Connection.privateEventName);
                    }
                    else if (request.variantId != null)
                        process.Variant_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), Connection.privateEventName);
                    else
                        process.As_IS_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), Connection.privateEventName);


                    lock (Processes)
                    {
                        Processes.Add(process);
                    }

                    if (process.As_IS_Request != null)
                    {
                        SendStatusMessage("GetDataRequest sent for AS-IS");
                        SendDataModuleMessage(process.As_IS_Request);
                    }

                    if (process.Variant_Request != null)
                    {
                        SendStatusMessage("GetDataRequest sent for Variant");
                        SendDataModuleMessage(process.Variant_Request);
                    }
                }

            }
        }

        private void HandleStartModuleRequest2(StartModuleRequest request)
        {
            if (request != null & request.moduleId == ModuleId)
            {
                SendStatusMessage("Handles StartModuleRequest");

                #region
                //    //Case: Hovsjo id mapping
                //if (request.caseId == "56cd5745317e88872c28ec12" & ModuleId == "SP_LCA_v4.0")
                //    {
                //        switch (request.variantId)
                //        {
                //            case null: //AsIS    
                //                SendStartModuleResponse(request, ModuleStatus.Success, "Ok", 0);
                //                break;
                //            case "56cd61a1317e88872c28ec16": // Alternative 1A
                //                var Data = ReadInputFile("Modules_input_output_LCA_20160219.xlsx", 
                //                      "DIMOSIM_RESULTS_AsIs.xlsx",
                //                      "DIMOSIM_RESULTS_1A.xlsx");
                //                var process = new ModuleProcess(request);
                //                process.As_IS_Data = Data;
                //                process.Variant_Data = Data;
                //                process.Variant_Request = new GetDataRequest(request,Guid.NewGuid().ToString());
                //                process.As_IS_Request = new GetDataRequest(request, Guid.NewGuid().ToString());
                //                CalculateResult(process);
                //                break;
                //            case "56d469d64a93972f1f61ad4a": // Alternative 2
                //            case "56d469e64a93972f1f61ad4b": // Alternative 3
                //            case "56d46a0c4a93972f1f61ad4c": // Alternative 4A
                //            case "56d46a354a93972f1f61ad4d": // Alternative 4B
                //            case "56d6faae4a93972f1f61ae7c": // Alternative Green
                //                break;
                //            default:
                //                break;
                //        }
                //    }
                //else if (request.caseId == "56cd5745317e88872c28ec12" & ModuleId == "Stockholm_Green_Area_Factor")
                //{
                //    var process = new ModuleProcess(request);
                //    switch (request.variantId)
                //    {
                //        case null: //AsIS    
                //        case "56cd61a1317e88872c28ec16": // Alternative 1A
                //        case "56d469d64a93972f1f61ad4a": // Alternative 2
                //        case "56d469e64a93972f1f61ad4b": // Alternative 3
                //        case "56d46a0c4a93972f1f61ad4c": // Alternative 4A
                //        case "56d46a354a93972f1f61ad4d": // Alternative 4B
                //            process.As_IS_Request = new GetDataRequest(request, Guid.NewGuid().ToString());
                //            process.As_IS_Data = new Dictionary<string, object>(); 
                //            CalculateResult(process);
                //            break;
                //        case "56d6faae4a93972f1f61ae7c": // Alternative Green
                //            process.Variant_Request = new GetDataRequest(request, Guid.NewGuid().ToString());
                //            process.Variant_Data = new Dictionary<string, object>();
                //            CalculateResult(process);
                //            break;
                //        default:
                //            break;
                //    }
                //}
                //else 
                #endregion
                if (ModuleId == "Stockholm_Green_Area_Factor")
                {
                    SendStartModuleResponse(request, ModuleStatus.Processing, "Accessing Data");
                    var process = new ModuleProcess(request);
                    if (request.variantId == null)
                    {
                        process.As_IS_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), Connection.privateEventName);
                        process.As_IS_Data = new Dictionary<string, object>();
                        CalculateResult(process);
                    }
                    else
                    {
                        process.Variant_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), Connection.privateEventName);
                        process.Variant_Data = new Dictionary<string, object>();
                        CalculateResult(process);
                    }

                }
                else
                {
                    SendStartModuleResponse(request, ModuleStatus.Processing, "Accessing Data");
                    if (request.variantId == null)
                    {
                        SendStartModuleResponse(request, ModuleStatus.Success, "Ok", 0);
                    }
                    else
                    {
                        Dictionary<string, object> Data = GenerateLCA_Variant_1A();
                        var process = new ModuleProcess(request);
                        process.As_IS_Data = Data;
                        process.Variant_Data = Data;
                        process.Variant_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), Connection.privateEventName);
                        process.As_IS_Request = new GetDataRequest(request, Guid.NewGuid().ToString(), Connection.privateEventName);
                        CalculateResult(process);
                    }
                }


                return;
            }
        }


        private Dictionary<string, object> GenerateLCA_Variant_1A()
        {
            Dictionary<string, object> Data;

            if (File.Exists("Modules_input_output_LCA_20160219.xlsx") &
                File.Exists("DIMOSIM_RESULTS_AsIs.xlsx") &
                File.Exists("DIMOSIM_RESULTS_1A.xlsx"))
                Data = ReadInputFile("Modules_input_output_LCA_20160219.xlsx",
                     "DIMOSIM_RESULTS_AsIs.xlsx",
                     "DIMOSIM_RESULTS_1A.xlsx");
            else
                Data = ReadInputFile();

            return Data;
        }

        private Dictionary<string, object> ReadInputFile(
            string file1 = @"C:\Users\johannesma\Documents\Dokument\!Ecodistr-ICT\Hovsjö\Modules_input_output_LCA_20160219.xlsx",
            string file2 = @"C:\Users\johannesma\Documents\Dokument\!Ecodistr-ICT\Hovsjö\DIMOSIM_RESULTS_AsIs.xlsx",
            string file3 = @"C:\Users\johannesma\Documents\Dokument\!Ecodistr-ICT\Hovsjö\DIMOSIM_RESULTS_1A.xlsx",
            string sheet = "Input Alt 1A",
            string sheet2 = "BUILDINGS RESULTS")
        {
            file1 = Path.GetFullPath(file1);
            file2 = Path.GetFullPath(file2);
            file3 = Path.GetFullPath(file3);


            Dictionary<string, object> data = new Dictionary<string, object>();
            var efile1 = new CExcel();
            var efile2 = new CExcel();
            var efile3 = new CExcel();
            try
            {
                GeoValue buildings = new GeoValue();
                buildings.features = new Features();
                data.Add("heat_source_before", 6); //"direct_electricity"
                data.Add("calculation_period", 30);
                data.Add("electricity_mix", 1); //"Sweden"
                data.Add("gwp_district", 83);
                data.Add("peu_district", 0.11);


                if (efile1.OpenWorkBook(file1) & efile2.OpenWorkBook(file2) & efile3.OpenWorkBook(file3))
                {
                    for (int i = 0; i < 105; ++i)
                    {
                        var building = new Ecodistrict.Messaging.Feature();
                        building.properties = new Dictionary<string, object>();
                        building.properties.Add("gml_id", efile2.GetCellValue(sheet2, i + 3, 3));

                        #region Get Data
                        //Change
                        // M + N  +  P + Q 
                        var ahd_before = Convert.ToDouble(efile2.GetCellValue(sheet2, i + 3, 13)) +
                                         Convert.ToDouble(efile2.GetCellValue(sheet2, i + 3, 14)) +
                                         Convert.ToDouble(efile2.GetCellValue(sheet2, i + 3, 16)) +
                                         Convert.ToDouble(efile2.GetCellValue(sheet2, i + 3, 17));
                        var ahd_after = Convert.ToDouble(efile3.GetCellValue(sheet2, i + 3, 13)) +
                                        Convert.ToDouble(efile3.GetCellValue(sheet2, i + 3, 14)) +
                                        Convert.ToDouble(efile3.GetCellValue(sheet2, i + 3, 16)) +
                                        Convert.ToDouble(efile3.GetCellValue(sheet2, i + 3, 17));
                        var ahdChange = ahd_after - ahd_before;
                        building.properties.Add("change_in_ahd_due_to_renovations_of_bshell_ventilation_pump", ahdChange);


                        var aed_before = Convert.ToDouble(efile2.GetCellValue(sheet2, i + 3, 18));
                        var aed_after = Convert.ToDouble(efile3.GetCellValue(sheet2, i + 3, 18));
                        var aedChange = aed_before - aed_after;
                        building.properties.Add("change_in_aed_due_to_renovations_of_bshell_ventilation_pump", aedChange);


                        var aed_fc_before = Convert.ToDouble(efile2.GetCellValue(sheet2, i + 3, 15));
                        var aed_fc_after = Convert.ToDouble(efile3.GetCellValue(sheet2, i + 3, 15));
                        var aed_fcChange = aed_fc_before - aed_fc_after;
                        building.properties.Add("change_in_aed_fc_due_to_renovations_of_bshell_ventilation_pump", aed_fcChange);

                        int offset = 11;
                        object datai;
                        //Insulation 1                        
                        datai = efile1.GetCellValue(sheet, 27, i * 2 + offset);
                        if (datai != null)
                        {
                            building.properties.Add("change_insulation_material_1", true);
                            building.properties.Add("insulation_material_1_life_of_product", datai);
                            building.properties.Add("insulation_material_1_type_of_insulation", efile1.GetCellValue(sheet, 28, i * 2 + offset));
                            building.properties.Add("insulation_material_1_amount_of_new_insulation_material", efile1.GetCellValue(sheet, 29, i * 2 + offset));
                        }

                        //Windows
                        datai = efile1.GetCellValue(sheet, 39, i * 2 + offset);
                        if (datai != null)
                        {
                            building.properties.Add("change_windows", true);
                            building.properties.Add("windows_life_of_product", datai);
                            building.properties.Add("windows_type_windows", efile1.GetCellValue(sheet, 40, i * 2 + offset));
                            building.properties.Add("windows_area_of_new_windows", efile1.GetCellValue(sheet, 41, i * 2 + offset));
                        }

                        //Ventilation System - Ventilation Ducts
                        datai = efile1.GetCellValue(sheet, 47, i * 2 + offset);
                        if (datai != null)
                        {
                            building.properties.Add("change_ventilation_ducts", true);
                            building.properties.Add("ventilation_ducts_life_of_product", datai);
                            building.properties.Add("ventilation_ducts_type_of_material", efile1.GetCellValue(sheet, 48, i * 2 + offset));
                            building.properties.Add("ventilation_ducts_weight_of_ventilation_ducts", efile1.GetCellValue(sheet, 49, i * 2 + offset));
                        }

                        //Ventilation System - Airflow assembly
                        datai = efile1.GetCellValue(sheet, 51, i * 2 + offset);
                        if (datai != null)
                        {
                            building.properties.Add("change_airflow_assembly", true);
                            building.properties.Add("airflow_assembly_life_of_product", datai);
                            building.properties.Add("airflow_assembly_type_of_airflow_assembly", efile1.GetCellValue(sheet, 52, i * 2 + offset));
                            building.properties.Add("airflow_assembly_design_airflow_exhaust_air", 60 * 60 * Convert.ToDouble(efile1.GetCellValue(sheet, 53, i * 2 + offset)));
                        }

                        //Water taps
                        datai = efile1.GetCellValue(sheet, 62, i * 2 + offset);
                        if (datai != null)
                        {
                            building.properties.Add("change_water_taps", true);
                            building.properties.Add("water_taps_life_of_product", datai);
                            building.properties.Add("number_of_taps", efile1.GetCellValue(sheet, 63, i * 2 + offset));
                        }
                        #endregion

                        buildings.features.Add(building);
                    }



                }
                else
                    SendStatusMessage("Could not opend excelfile");

                data.Add("buildings", buildings);


            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "OpenWorkbook", exception: ex);
            }
            finally
            {
                efile1.CloseExcel();
                efile1.CloseWorkBook();
                efile2.CloseExcel();
                efile2.CloseWorkBook();
                efile3.CloseExcel();
                efile3.CloseWorkBook();
            }

            return data;
        }

        private void HandleGetDataResponse(GetDataResponse response)
        {
            lock (Processes)
            {
                foreach (ModuleProcess process in Processes)
                {
                    bool calculate = false;
                    bool matchFound = false;

                    if (response.variantId == null)
                    {
                        if (process.As_IS_Request.calculationId == response.calculationId)
                        {
                            matchFound = true;
                            process.As_IS_Data = response.data;
                            if ((process.Request.variantId == null) |
                                (useBothVariantAndAsISForVariant & process.Variant_Data != null))
                                calculate = true;
                        }
                    }

                    if (response.variantId != null)
                    {

                        if (process.Variant_Request.calculationId == response.calculationId)  //TODO CalculationId not returned
                        {
                            matchFound = true;
                            process.Variant_Data = response.data;
                            if ((useBothVariantAndAsISForVariant & process.As_IS_Data != null) |
                                !useBothVariantAndAsISForVariant)
                                calculate = true;
                        }
                    }

                    if (calculate)
                    {
                        SendStartModuleResponse(process.Request, ModuleStatus.Processing, "Received message from data module");
                        CalculateResult(process);

                        Processes.Remove(process);
                    }

                    if (matchFound)
                        return;

                }
            }
        }

        #endregion

        #region Send Hub Messages

        void SendMessage(IMessage message)
        {
            var str = Serialize.ToJsonString(message);
            Publish(str);
        }

        void SendDataModuleMessage(IMessage message)
        {
            var str = Serialize.ToJsonString(message);
            PublishData(str);
        }

        /// <summary>
        /// Returns a GetModuleResponse to the dashboard
        /// </summary>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private bool SendGetModulesResponse()
        {
            try
            {
                GetModulesResponse gmRes = new GetModulesResponse(ModuleName, ModuleId, Description, KpiList);
                SendMessage(gmRes);
                SendStatusMessage("GetModulesResponse sent");
                return true;
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendGetModulesResponse", exception: ex);
                return false;
            }

        }

        /// <summary>
        /// Sends a SelectModuleResponse to the dashboard
        /// </summary>
        /// <param name="request"></param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private bool SendSelectModuleResponse(SelectModuleRequest request)
        {
            try
            {
                var variantId = request.variantId;
                var caseId = request.caseId;
                var kpiId = request.kpiId;
                var smResponse = new SelectModuleResponse(ModuleId, variantId, caseId, kpiId);
                SendMessage(smResponse);
                SendStatusMessage("SelectModuleResponse sent");
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendSelectModuleResponse", exception: ex);
                return false;
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="request"></param>
        /// <param name="status"></param>
        /// <param name="info"></param>
        /// <returns></returns>
        private bool SendStartModuleResponse(StartModuleRequest request, ModuleStatus status = ModuleStatus.Processing, string info = "", double? kpiValue = null)
        {
            try
            {
                var smr = new StartModuleResponse(request, status, info, kpiValue);
                SendMessage(smr);
                SendStatusMessage(string.Format("StartModuleResponse {0} sent", status.ToString()));
            }
            catch (Exception ex)
            {
                try
                {
                    string errInfo = "Internal module error: contact developer";
                    var smr = new StartModuleResponse(request, ModuleStatus.Failed, errInfo);
                    SendMessage(smr);
                    SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-StartmoduleResponse", exception: ex);
                }
                catch
                {
                    SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-StartmoduleResponse", exception: ex);
                }
                return false;
            }

            return true;
        }

        #endregion

        #region Internal message
        protected bool CheckAndReportBuildingProp(ModuleProcess process, Feature building, string key)
        {
            if (!building.properties.ContainsKey(key))
            {
                string buildingIdKey = "attr_gml_id";
                if (building.properties.ContainsKey(buildingIdKey))
                    process.CalcMessage = String.Format("Property {0} missing in building {1}", key, building.properties[buildingIdKey]);
                else
                    process.CalcMessage = String.Format("Property {0} missing, building id not set", key);

                return false;
            }

            return true;
        }
        protected bool CheckAndReportBuildingProp(ModuleProcess process, Dictionary<string, object> buildingProps, string key)
        {
            if (!buildingProps.ContainsKey(key))
            {
                string buildingIdKey = "attr_gml_id";
                if (buildingProps.ContainsKey(buildingIdKey))
                    process.CalcMessage = String.Format("Property {0} missing in building {1}", key, buildingProps[buildingIdKey]);
                else
                    process.CalcMessage = String.Format("Property {0} missing, building id not set", key);

                return false;
            }

            return true;
        }
        protected bool CheckAndReportDistrictProp(ModuleProcess process, Dictionary<string, object> distrProps, string key)
        {
            if (distrProps == null)
            {
                process.CalcMessage = String.Format("Property {0} is missing in district data", key);
                return false;
            }
            else if (!distrProps.ContainsKey(key))
            {
                process.CalcMessage = String.Format("Property {0} is missing in district data", key);
                return false;
            }
            else if (distrProps[key] == null)
            {
                process.CalcMessage = String.Format("Property {0} is missing in district data", key);
                return false;
            }

            return true;
        }
        #endregion

        #region Write to Console
        /// <summary>
        /// When set (default) status messages is only sent when the module sends something to the dashboard
        /// not when it receives something from the dashboard.
        /// </summary>
        protected bool ShowOnlyOwnStatus { get; set; }
        /// <summary>
        /// The error event that could be subscribed to
        /// </summary>
        public event ErrorEventHandler ErrorRaised;
        /// <summary>
        /// The status message event that could be subscribed to
        /// </summary>
        public event StatusEventHandler StatusMessage;

        protected void SendStatusMessage(string message)
        {
            if (StatusMessage != null)
            {
                var e = new StatusEventArg { StatusMessage = message };
                StatusMessage(this, e);
            }
        }

        protected void SendErrorMessage(string message, string sourceFunction, Exception exception = null)
        {
            if (ErrorRaised != null)
            {
                var e = new ErrorMessageEventArg { Message = message, SourceFunction = sourceFunction, Exception = exception };
                ErrorRaised(this, e);
            }
        }

        #endregion

        #region Kpi Calculation
        protected abstract bool CalculateKpi(ModuleProcess process, CExcel exls, out Ecodistrict.Messaging.Data.Output output, out Ecodistrict.Messaging.Data.OutputDetailed outputDetailed);

        private bool CalculateResult(ModuleProcess process)
        {
            Ecodistrict.Messaging.Data.Output output = null;
            Ecodistrict.Messaging.Data.OutputDetailed outputDetailed = null;
            try
            {
                if (File.Exists(WorkBookPath))
                {
                    if (ExcelApplikation.OpenWorkBook(WorkBookPath))
                    {
                        //Calculate KPI
                        if (CalculateKpi(process, ExcelApplikation, out output, out outputDetailed))
                        {
                            //Send Detailed Result to DB 
                            if (outputDetailed != null)
                            {
                                var setRes = new SetKpiResultRequest(ModuleId, process.Request.caseId, process.Request.variantId, outputDetailed);
                                var str = Serialize.ToJsonString(setRes);
                                PublishData(str);
                                SendStatusMessage("SetKpiResultRequest sent");
                            }
                            // Send Mean Kpi to dashboard
                            SendStartModuleResponse(process.Request, ModuleStatus.Success, "Ok", output.KpiValue);
                            //SendStatusMessage(str);

                        }
                        else
                        {
                            SendStartModuleResponse(process.Request, ModuleStatus.Failed, process.CalcMessage);
                            SendErrorMessage("Could not calculate kpi: " + process.CalcMessage, "CalculateKpi");
                            return false;
                        }
                    }
                }
                else
                {
                    SendStartModuleResponse(process.Request, ModuleStatus.Failed, "Dependent calculation module not found, contact developer");
                    SendErrorMessage(string.Format("Excelfile <{0}> not found", WorkBookPath), sourceFunction: "SendModuleResult-FileNotFound");
                    return false;
                }

            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-KalkKpi", exception: ex);
                SendStartModuleResponse(process.Request, ModuleStatus.Failed, "Internal module error: contact developer");

                return false;
            }
            finally
            {
                ExcelApplikation.CloseWorkBook();
            }

            return true;

        }

        #endregion

    }


    public class ModuleProcess : IEquatable<ModuleProcess>
    {
        public string KpiId
        {
            get
            {
                if (Request != null)
                    return Request.kpiId;

                else return "";
            }
        }
        public ModuleProcess(StartModuleRequest request, double timelimit = 2000)
        {
            this.DateTime = DateTime.Now;
            this.Guid = Guid.NewGuid();
            this.Request = request;
            _timeLimit = timelimit;
        }
        private double _timeLimit;
        public DateTime DateTime { get; private set; }
        public Guid Guid { get; private set; }
        public StartModuleRequest Request { get; private set; }
        public Dictionary<string, object> As_IS_Data { get; set; }
        public Dictionary<string, object> Variant_Data { get; set; }
        public GetDataRequest As_IS_Request { get; set; }
        public GetDataRequest Variant_Request { get; set; }
        public bool Equals(ModuleProcess process)
        {
            return this.Guid == process.Guid;
        }

        public bool TimerExpired
        {
            get
            {
                return ((DateTime.Now - DateTime).TotalMilliseconds > _timeLimit);
            }
        }

        public Dictionary<string, object> CurrentData
        {
            get
            {
                if (!IsAsIS) //Its a variant
                    return Variant_Data;
                else
                    return As_IS_Data;
            }
        }

        public bool IsAsIS
        {
            get
            {
                return Variant_Data == null;
            }
        }

        public string CalcMessage { get; set; }

    }
}

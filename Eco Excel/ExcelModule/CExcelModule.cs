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
        /// <summary>
        /// The name of the IMB subscription the application uses. Read from the configuration file
        /// </summary>
        protected  string SubScribedEventName { get; set; }
        /// <summary>
        /// The name the application uses when sending back information to the dashboard
        /// </summary>
        protected string PublishedEventName { get; set; }

        public  TConnection Connection { get; set; }
        public TEventEntry SubscribedEvent { get; set; }
        public TEventEntry PublishedEvent { get; set; }
        /// <summary>
        /// The error event that could be subscribed to
        /// </summary>
        public event ErrorEventHandler ErrorRaised;
        /// <summary>
        /// The status message event that could be subscribed to
        /// </summary>
        public event StatusEventHandler StatusMessage;

        /// <summary>
        /// When set (default) statusmessages is only sent when the mudule sends something to the dashboard
        /// not when it receives something from the dashboard.
        /// </summary>
        protected bool ShowOnlyOwnStatus { get; set; }
            
        private CExcel ExcelApplikation { get; set; }
        
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
        /// Name of the module owner/responsable
        /// </summary>
        protected string UserName { get; set; }

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
        /// A list of strings with Kpis that the ExcelFile can calculate.
        /// </summary>
        protected List<string> KpiList { get; set; }
        /// <summary>
        /// Description of the module.
        /// </summary>
        protected string Description { get; set; }

        /// <summary>
        /// Creates a new CExcel instance that in turn creates a new instance of Excel.Application
        /// </summary>
        protected CExcelModule()
        {
            try
            {
                ShowOnlyOwnStatus = true;
                ExcelApplikation = new CExcel();


                // Create a Timer object that knows to call our TimerCallback
                // method once every n milliseconds.
                timer = new Timer(TestConnection);

            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "CExcel Constructor", exception: ex);
            }
        }
        Timer timer;
        private void TestConnection(Object o)
        {
            try
            {
                if (SendGetModulesResponse())
                    return;

                timer.Change(Timeout.Infinite, Timeout.Infinite);

                SendStatusMessage("IMB connection lost");
                if (Connection != null)
                    Connection.Dispose();
                Connection = null;

                SendStatusMessage("Try to reconnect to IMB..");
                if (!ReConnect(20))
                    SignalConnectionLost();
            }
            catch (System.Exception ex)
            {
                SendErrorMessage("Connection lost", "", ex);
                return;
            }
        }
        public event EventHandler ConnectionLost;
        public void SignalConnectionLost()
        {            
            EventHandler handler = this.ConnectionLost;

            if (handler != null)
            {
                handler(this, new EventArgs());
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

        //IMB hub info
        string aCertFile = "client-eco-district.pfx";
        string aCertFilePassword = "&8dh48klosaxu90OKH";
        string aRootCertFile = "root-ca-imb.crt";
        string aPrefix = "ecodistrict";
        protected string RemoteHost { get; set; }

        /// <summary>
        /// Close any opened Excel dokument and Excel.Application as a preparation for closing down.
        /// If anything more has to be done it can be overrided.
        /// </summary>
        public virtual void Close()
        {
            try
            {
                timer.Change(Timeout.Infinite, Timeout.Infinite);

                if (ExcelApplikation != null)
                    ExcelApplikation.CloseExcel();

                ExcelApplikation = null;

                if (Connected)
                {
                    //Close connection
                    Connection.close();
                    // reset event handler for change object on subscribedEvent
                    if (SubscribedEvent!=null)
                        SubscribedEvent.onString -= SubscribedEvent_onString;
                }
            }
            catch (Exception ex)
            {
                // reset event handler for change object on subscribedEvent
                if (SubscribedEvent != null)
                    SubscribedEvent.onString -= SubscribedEvent_onString;
                SendErrorMessage(message: ex.Message, sourceFunction: "Close", exception: ex);
            }

            if (Connection != null)
                Connection.Dispose();
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
	                Connection = new TTLSConnection(aCertFile, aCertFilePassword, aRootCertFile, ModuleName, UserId, aPrefix, RemoteHost);
	
	                if (Connection.connected)
	                {
	                    SubscribedEvent = Connection.subscribe(SubScribedEventName);
	                    PublishedEvent = Connection.publish(PublishedEventName);
	
	                    // set event handler for change object on subscribedEvent
	                    SubscribedEvent.onString += SubscribedEvent_onString;

                        SendStatusMessage("Connected to IMB-hub..");
                        res = true;
                        
                        timer.Change(6000, 3 * 60000);
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
               
        #region Events
        void SubscribedEvent_onString(TEventEntry aEventEntry, string msg)
        {
            try
            {
                //Deserialize only header to prevent unnecessary deserialization if the message was not meant for this module 
                Request header = Deserialize<Request>.JsonString(msg); 

                if (header != null)
                {

                    if (header is GetModulesRequest)
                    {
                        SendStatusMessage("GetModulesRequest received");
                        if (!SendGetModulesResponse())
                            SendErrorMessage(message: "could not send getModulesResponse", sourceFunction: "SubscribedEvent_OnNormalEvent");
                    }
                    else if (ModuleId == header.moduleId)
                    {
                        IMessage iMessage = Deserialize<IMessage>.JsonString(msg);

                        if (iMessage is SelectModuleRequest)
	                    {
	                        if (!ShowOnlyOwnStatus)
	                            SendStatusMessage("SelectModuleRequest received");
	                        
	                        SendStatusMessage("Handles SelectModuleRequest");
                            SendSelectModuleResponse(iMessage as SelectModuleRequest);
                                //if (!SendSelectModuleResponse(iMessage as SelectModuleRequest))
                                //    SendErrorMessage(message: "could not send SelectModulesResponse", sourceFunction: "SubscribedEvent_OnNormalEvent");
	                    }
                        else if (iMessage is StartModuleRequest)
	                    {
	                        if (!ShowOnlyOwnStatus)
	                            SendStatusMessage("StartModuleRequest received");
	                        
	                        SendStatusMessage("Handles StartModuleRequest");
                            SendModuleResult(iMessage as StartModuleRequest);
                                //if (!SendModuleResult(iMessage as StartModuleRequest))
                                //    SendErrorMessage(message: "could not send StartModulesesponse", sourceFunction: "SubscribedEvent_OnNormalEvent");
	                    }
                    }

                }
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SubscribedEvent_OnNormalEvent", exception: ex);
                //TMP - Store data locally
                string path = Path.GetDirectoryName(this.WorkBookPath);
                System.IO.File.WriteAllText(String.Format(@"{0}/{1}{2} {3}.json", path, this.UserName, "Error - Message", DateTime.Now.ToString("yyyy/MM/dd HH.mm.ss")), msg);
                //
            }
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

        protected virtual void Init_IMB(string IMB_config_path)
        {
            try
            {
                var serializer = new YamlSerializer();
                var imb_settings = serializer.DeserializeFromFile(IMB_config_path, typeof(IMB_Settings))[0];

                this.RemoteHost = ((IMB_Settings)imb_settings).remoteHost;
                this.SubScribedEventName = ((IMB_Settings)imb_settings).subScribedEventName;
                this.PublishedEventName = ((IMB_Settings)imb_settings).publishedEventName;
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
                return true;
            }
            catch (Exception ex)
            {
                CExcelModule_ErrorRaised(this, ex);
                return false;
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

        protected virtual InputSpecification GetInputSpecification(string kpiId)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// This function is to be inherited. It receives all indata parameters from the dashboard as a dictionary, the name of 
        /// the Kpi that it should calculate and the Excel object that should be used for the calculations. 
        /// </summary>
        /// <param name="indata">indata from the dashboard</param>
        /// <param name="kpiId">The name of the Kpi that is to be calculated</param>
        /// <param name="exls">Excel object</param>
        /// <returns>A output object that can be serialized and sent to th dashboard</returns>
        protected virtual Ecodistrict.Messaging.Output.Outputs CalculateKpi(Dictionary<string,Input> indata,string kpiId, CExcel exls)
        {
            throw new NotImplementedException();
            
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

                var str = Serialize.ToJsonString(gmRes);
                Publish(str);
                //PublishedEvent.signalString(str);
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
                var kpiId = request.kpiId;
                var smResponse = new SelectModuleResponse(ModuleId, variantId, kpiId, GetInputSpecification(kpiId));

                var str = Serialize.ToJsonString(smResponse);
                Publish(str);
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
        /// Handles the StartModuleRequest from the dashboard.<br/> 
        /// Starts with sending a StartModuleResponse processing message to the dashboard. After that
        /// the Excel document file is opened and the function calculateKpi <see cref="CalculateKpi"/> function is called.
        /// The result from the calculation is an ouputs object that is sent back to the dashboard.
        /// The last message sent to the dashbord is a StartModuleResponse Success processing.<br/>
        /// At last the >Excel document is closed. No changes i 
        /// If anything goes wrong with the calculations a StartModuleResponse Failure is sent to the dashboard
        /// </summary>
        /// <param name="request">The requst sent from the dashboard.</param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private bool SendModuleResult(StartModuleRequest request)
        {
            try
            {
                var smr = new StartModuleResponse(ModuleId, request.variantId, request.userId, request.kpiId, ModuleStatus.Processing);
                var str = Serialize.ToJsonString(smr);
                Publish(str);
                SendStatusMessage("StartModuleResponse processing sent"); 
            }
            catch (Exception ex)
            {
                try
                {
                    var smr = new StartModuleResponse(ModuleId, request.variantId, request.userId, request.kpiId, ModuleStatus.Failed);
                    var str = Serialize.ToJsonString(smr);
                    Publish(str);
                    SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-StartmoduleResponse", exception: ex);
                }
                catch 
                {
                    SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-StartmoduleResponse", exception: ex);
                }
                return false;
            }

            Ecodistrict.Messaging.Output.Outputs outputs = null;


            try
            {
                if (File.Exists(WorkBookPath))
                {
                    if (ExcelApplikation.OpenWorkBook(WorkBookPath))
                    {
                        outputs = CalculateKpi(request.inputs, request.kpiId, ExcelApplikation);
                        //TMP - Store data locally
                        string dataStr = Serialize.ToJsonString(request);
                        string path = Path.GetDirectoryName(this.WorkBookPath);
                        System.IO.File.WriteAllText(String.Format(@"{0}/{1}{2} {3}.json", path, this.UserName, "Message - StartModuleRequest ", DateTime.Now.ToString("yyyy/MM/dd HH.mm.ss")), dataStr);
                        //
                    }
                }
                else
                {
                    var smr = new StartModuleResponse(ModuleId, request.variantId, request.userId, request.kpiId, ModuleStatus.Failed);
                    var str = Serialize.ToJsonString(smr);
                    Publish(str);
                    SendErrorMessage(string.Format("Excelfile <{0}> not found", WorkBookPath), sourceFunction: "SendModuleResult-FileNotFound");
                    return false;
                }

            }
            catch (Exception ex)
            {

                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-KalkKpi", exception: ex);

                var stmResp2 = new StartModuleResponse(ModuleId, request.variantId, request.userId, request.kpiId, ModuleStatus.Failed);
                var str = Serialize.ToJsonString(stmResp2);
                Publish(str);
                SendStatusMessage("StartModuleResponse Failed sent");

                //TMP - Store data locally
                string dataStr = Serialize.ToJsonString(request);
                string path = Path.GetDirectoryName(this.WorkBookPath);
                System.IO.File.WriteAllText(String.Format(@"{0}/{1}{2} {3}.json", path, this.UserName, "Error - StartModuleRequest ", DateTime.Now.ToString("yyyy/MM/dd HH.mm.ss")), dataStr);
                //
               
                return false;
            }
            finally
            {
               ExcelApplikation.CloseWorkBook();
            }

            try
            {
                ModuleResult result = new ModuleResult(ModuleId, request.variantId, request.userId, request.kpiId, outputs);
                var str = Serialize.ToJsonString(result);

                Publish(str);
                SendStatusMessage("ModuleResult sent");

                //TMP - Store data locally
                string dataStr = Serialize.ToJsonString(request);
                string path = Path.GetDirectoryName(this.WorkBookPath);
                System.IO.File.WriteAllText(String.Format(@"{0}/{1}{2} {3}.json", path, this.UserName, "Message - ModuleResult", DateTime.Now.ToString("yyyy/MM/dd HH.mm.ss")), str);
                //
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-SendModuleResult", exception: ex);
                return false;
            }
            
            return true;

        }

        private void SendStatusMessage(string message)
        {
            if (StatusMessage != null)
            {
                var e = new StatusEventArg {StatusMessage = message};
                StatusMessage(this, e);
            }
        }

        private void SendErrorMessage(string message, string sourceFunction, Exception exception=null)
        {
            if (ErrorRaised != null)
            {
                var e = new ErrorMessageEventArg {Message = message,SourceFunction =sourceFunction, Exception = exception};
                ErrorRaised(this, e);
            }        
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

    }
}

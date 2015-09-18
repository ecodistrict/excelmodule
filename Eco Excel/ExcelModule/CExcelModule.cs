using System;
using System.Collections.Generic;
using System.IO;
using IMB;
using Ecodistrict.Messaging;

namespace Ecodistrict.Excel
{
    /// <summary>
    /// Eventhandler used for Error reporting
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

        private  TConnection Connection { get; set; } 
        private  TEventEntry SubscribedEvent { get; set; }
        private  TEventEntry PublishedEvent { get; set; }
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

        /// <summary>
        /// The complete path to the Excedocument document file that the module is going to use (*.xls, *.xlsx)
        /// </summary>
        protected string workBookPath { get; set; }
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


        string aCertFile = "client-eco-district.pfx";
        string aCertFilePassword = "&8dh48klosaxu90OKH";
        string aRootCertFile = "root-ca-imb.crt";

        string aPrefix = "ecodistrict";
        string aRemoteHost = "192.168.239.138";//"vps17642.public.cloudvps.com";//"192.168.239.138"

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

                if (Connection.connected)
                    Connection.close();
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "Close", exception: ex);
            }
        }
        /// <summary>
        /// Connects to hub/Server, prepares the publish event and starts subscription to dashboard events
        /// </summary>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        public virtual bool ConnectToServer()
        {
            bool res = true;
            
            try
            {
                Connection = new TTLSConnection(aCertFile, aCertFilePassword, aRootCertFile, UserName, UserId, aPrefix, aRemoteHost);

                if (Connection.connected)
                {
                    SubscribedEvent = Connection.subscribe(SubScribedEventName);
                    PublishedEvent = Connection.publish(PublishedEventName);

                    // set event handler for change object on subscribedEvent
                    SubscribedEvent.onString += SubscribedEvent_onString;
                }
                else
                {
                    //Console.WriteLine("## NOT connected");
                    res = false;
                }

            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "ConnectToServer", exception: ex);
                res = false;
            }
            return res;

        }

        void SubscribedEvent_onString(TEventEntry aEventEntry, string msg)
        {
            try
            {

                IMessage iMessage = Deserialize.JsonString(msg);

                if (iMessage != null)
                {
                    if (iMessage is GetModulesRequest)
                    {
                        SendStatusMessage("GetModulesRequest received");
                        if (!SendGetModulesResponse())
                            SendErrorMessage(message: "could not send getModulesResponse", sourceFunction: "SubscribedEvent_OnNormalEvent");
                    }
                    else if (iMessage is SelectModuleRequest)
                    {
                        if (!ShowOnlyOwnStatus)
                            SendStatusMessage("SelectModuleRequest received");

                        var smr = iMessage as SelectModuleRequest;
                        if (ModuleId == smr.moduleId)
                        {
                            SendStatusMessage("Handles SelectModuleRequest");
                            if (!SendSelectModuleResponse(smr))
                                SendErrorMessage(message: "could not send SelectModulesResponse", sourceFunction: "SubscribedEvent_OnNormalEvent");
                        }
                    }
                    else if (iMessage is StartModuleRequest)
                    {
                        if (!ShowOnlyOwnStatus)
                            SendStatusMessage("StartModuleRequest received");

                        var smr = iMessage as StartModuleRequest;
                        if (ModuleId == smr.moduleId)
                        {
                            SendStatusMessage("Handles StartModuleRequest");
                            if (!SendModuleResult(smr))
                                SendErrorMessage(message: "could not send StartModulesesponse", sourceFunction: "SubscribedEvent_OnNormalEvent");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SubscribedEvent_OnNormalEvent", exception: ex);
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
                ExcelApplikation.OpenWorkBook(workBookPath);

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
        /// Returnes a GetModuleResponse to the dashboard
        /// </summary>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        private bool SendGetModulesResponse()
        {
            try
            {
               
                GetModulesResponse gmRes = new GetModulesResponse(ModuleName, ModuleId, Description, KpiList);

                var str = Serialize.ToJsonString(gmRes);
                PublishedEvent.signalString(str);
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
                PublishedEvent.signalString(str);
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
                PublishedEvent.signalString(str);
                SendStatusMessage("StartModuleResponse processing sent"); 
            }
            catch (Exception ex)
            {
                try
                {
                    var smr = new StartModuleResponse(ModuleId, request.variantId, request.userId, request.kpiId, ModuleStatus.Failed);
                    var str = Serialize.ToJsonString(smr);
                    PublishedEvent.signalString(str);
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
                if (File.Exists(workBookPath))
                {
                    if (ExcelApplikation.OpenWorkBook(workBookPath))
                    {
                        outputs = CalculateKpi(request.inputs, request.kpiId, ExcelApplikation);
                    }
                }
                else
                {
                    var smr = new StartModuleResponse(ModuleId, request.variantId, request.userId, request.kpiId, ModuleStatus.Failed);
                    var str = Serialize.ToJsonString(smr);
                    PublishedEvent.signalString(str);
                    SendErrorMessage(string.Format("Excelfile <{0}> not found", workBookPath), sourceFunction: "SendModuleResult-FileNotFound");
                    return false;
                }

            }
            catch (Exception ex)
            {

                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-KalkKpi", exception: ex);

                var stmResp2 = new StartModuleResponse(ModuleId, request.variantId, request.userId, request.kpiId, ModuleStatus.Failed);
                var str = Serialize.ToJsonString(stmResp2);
                PublishedEvent.signalString(str);
                SendStatusMessage("StartModuleResponse Failed sent"); 
               
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

                ////TMP - Store data locally
                //string path = Path.GetDirectoryName(this.workBookPath);
                //System.IO.File.WriteAllText(String.Format(@"{0}/{1}{2} {3}.json", path, this.UserName, " - ModuleResult", DateTime.Now.ToString("yyyy/MM/dd HH.mm.ss")), str);
                ////

                PublishedEvent.signalString(str);
                SendStatusMessage("ModuleResult sent"); 
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
    }
}

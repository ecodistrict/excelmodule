using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using IMB;
using IMB.ByteBuffers;
using Ecodistrict.Messaging;

namespace Ecodistrict.Excel
{
    public delegate void ErrorEventHandler(object sender, ErrorMessageEventArg e);
    public delegate void StatusEventHandler(object sender, StatusEventArg e);

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

        public event ErrorEventHandler ErrorRaised;
        public event StatusEventHandler StatusMessage;

        protected bool ShowOnlyOwnStatus { get; set; }
            
        private CExcel excelApplikation { get; set; }

        protected  string  ServerAdress { get; set; }
        protected  int Port { get; set; }
        protected int UserId { get; set; }
        protected string ModuleName { get; set; }
        protected string ModuleId { get; set; }
        protected string UserName { get; set; }
        protected string Federation { get; set; }
        protected string workBookPath { get; set; }
        protected List<string> KpiList { get; set; }
        protected string Description { get; set; }


        protected CExcelModule()
        {
            try
            {
                ShowOnlyOwnStatus = true;
                excelApplikation = new CExcel();
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "CExcel Constructor", exception: ex);
                return;
            }
        }

        ~CExcelModule()
        {
            Close();
        }

        public virtual void Close()
        {
            try
            {
                if (excelApplikation != null)
                    excelApplikation.CloseExcel();

                excelApplikation = null;

                if (Connection.Connected)
                    Connection.Close();
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "Close", exception: ex);
                return;

            }
        }

        public virtual bool ConnectToServer()
        {
            bool res = true;
            
            try
            {
                Connection = new TConnection(ServerAdress, Port, UserName, UserId, Federation);

                if (Connection.Connected)
                {
                    SubscribedEvent = Connection.Subscribe(SubScribedEventName);
                    PublishedEvent = Connection.Publish(PublishedEventName);

                    // set event handler for change object on subscribedEvent
                    SubscribedEvent.OnNormalEvent += SubscribedEvent_OnNormalEvent;
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

        protected virtual Ecodistrict.Messaging.InputSpecification GetInputSpecification(string kpiId)
        {
            throw new NotImplementedException();
        }

        protected virtual Ecodistrict.Messaging.Outputs CalculateKpi(Dictionary<string,object> indata,string kpiId, CExcel exls)
        {
            throw new NotImplementedException();
            
        }

        private void SubscribedEvent_OnNormalEvent(TEventEntry aEvent, TByteBuffer aPayload)
        {
            try
            {
                String msg = "";
                aPayload.Read(out msg);

                IMessage iMessage = Ecodistrict.Messaging.Deserialize.JsonString(msg);

                if (iMessage!=null)
                {
                    if (iMessage is Ecodistrict.Messaging.GetModulesRequest)
                    {
                        SendStatusMessage("GetModulesRequest received");
                        SendGetModulesResponse();
                    }
                    else if (iMessage is SelectModuleRequest)
                    {
                        if(!ShowOnlyOwnStatus)
                            SendStatusMessage("SelectModuleRequest received");

                        var smr = iMessage as SelectModuleRequest;
                        if (ModuleId == smr.moduleId)
                        {
                            SendStatusMessage("Handles SelectModuleRequest");
                            SendSelectModuleResponse(smr);
                        }
                    }
                    else if (iMessage is StartModuleRequest)
                    {
                        if (!ShowOnlyOwnStatus) 
                            SendStatusMessage("StartModuleRequest received");

                        var SMR = iMessage as StartModuleRequest;
                        if (ModuleId == SMR.moduleId)
                        {
                            SendStatusMessage("Handles StartModuleRequest");
                            SendModuleResult(SMR);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SubscribedEvent_OnNormalEvent", exception: ex);
                return;
            }
        }

        private bool SendGetModulesResponse()
        {
            try
            {
               
                GetModulesResponse gmRes = new GetModulesResponse(ModuleName, ModuleId, Description, KpiList);

                var str = Serialize.ToJsonString(gmRes);
                var payload = new TByteBuffer();
                payload.Prepare(str);
                payload.PrepareApply();
                payload.QWrite(str);
                PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);
                SendStatusMessage("GetModulesResponse sent");    
                return true;
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendGetModulesResponse", exception: ex);
                return false;
            }

        }

        private bool SendSelectModuleResponse(SelectModuleRequest request)
        {
            try
            {
                var variantId = request.variantId;
                var kpiId = request.kpiId;
                var smResponse = new SelectModuleResponse(ModuleId, variantId, kpiId, GetInputSpecification(kpiId));

                var str = Serialize.ToJsonString(smResponse);
                var payload = new TByteBuffer();
                payload.Prepare(str);
                payload.PrepareApply();
                payload.QWrite(str);
                PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);
                SendStatusMessage("SelectModuleResponse sent"); 
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendSelectModuleResponse", exception: ex);
                return false;
            }

            return true;
        }

        private bool SendModuleResult(StartModuleRequest request)
        {
            try
            {
                var smr = new StartModuleResponse(ModuleId, request.variantId, request.kpiId, ModuleStatus.Processing);
                var str = Serialize.ToJsonString(smr);
                var payload = new TByteBuffer();
                payload.Prepare(str);
                payload.PrepareApply();
                payload.QWrite(str);
                PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);
                SendStatusMessage("StartModuleResponse processing sent"); 
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-StartmoduleResponse",exception:ex);
                return false;
            }

            Outputs outputs = null;


            try
            {
                if (File.Exists(workBookPath))
                {
                if (excelApplikation.OpenWorkBook(workBookPath))
                {
                    outputs=CalculateKpi(request.inputData, request.kpiId, excelApplikation);
                }
                }
                else
                {
                    SendErrorMessage(string.Format("Excelfile <{0}> not found", workBookPath), sourceFunction: "SendModuleResult-FileNotFound");
                    return false;
                }

            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-KalkKpi", exception: ex);
                return false;
            }
            finally
            {
               excelApplikation.CloseWorkBook();
            }

            try
            {
                ModuleResult result = new ModuleResult(ModuleId, request.variantId, request.kpiId, outputs);
                var str = Ecodistrict.Messaging.Serialize.ToJsonString(result);
                var payload = new TByteBuffer();
                payload.Prepare(str);
                payload.PrepareApply();
                payload.QWrite(str);
                PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);
                SendStatusMessage("ModuleResult sent"); 
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-SendModuleResult", exception: ex);
                return false;
            }

            try
            {
                StartModuleResponse stmResp2 = new StartModuleResponse(ModuleId, request.variantId, request.kpiId,
                    ModuleStatus.Success);
                var str = Ecodistrict.Messaging.Serialize.ToJsonString(stmResp2);
                var payload = new TByteBuffer();
                payload.Prepare(str);
                payload.PrepareApply();
                payload.QWrite(str);
                PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);
                SendStatusMessage("StartModuleResponse success sent"); 
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "SendModuleResult-SendRSuccessResponse",exception:ex);
                return false;
            }

            return true;

        }

        private void SendStatusMessage(string message)
        {
            if (StatusMessage != null)
            {
                var e = new StatusEventArg() {StatusMessage = message};
                StatusMessage(this, e);
            }
        }

        private void SendErrorMessage(string message, string sourceFunction, Exception exception=null)
        {
            if (ErrorRaised != null)
            {
                var e = new ErrorMessageEventArg() {Message = message,SourceFunction =sourceFunction, Exception = exception};
                ErrorRaised(this, e);
            }        
        }
    }
}

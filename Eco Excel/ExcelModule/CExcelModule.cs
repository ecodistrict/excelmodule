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
    public delegate void ErrorEventHandler(object sender, EventArgs e);

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
                excelApplikation = new CExcel();
            }
            catch (Exception ex)
            {
                
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
                //Console.WriteLine(ex.Message);
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
            String msg = "";
            aPayload.Read(out msg);

            IMessage iMessage = Ecodistrict.Messaging.Deserialize.JsonString(msg);

            if (iMessage is Ecodistrict.Messaging.GetModulesRequest)
            {
                SendGetModulesResponse();
            }
            else if (iMessage is SelectModuleRequest)
            {
                var smr = iMessage as SelectModuleRequest;
                if (ModuleId == smr.moduleId)
                {
                    SendSelectModuleResponse(smr);
                }
            }
            else if (iMessage is StartModuleRequest)
            {
                var SMR = iMessage as StartModuleRequest;
                if (ModuleId == SMR.moduleId)
                {
                    SendModuleResult(SMR);
                }
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

                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        private bool SendSelectModuleResponse(SelectModuleRequest request)
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

            return true;
        }

        private bool SendModuleResult(StartModuleRequest request)
        {
            var smr = new StartModuleResponse(ModuleId, request.variantId, request.kpiId, ModuleStatus.Processing);

            var str = Serialize.ToJsonString(smr);
            var payload = new TByteBuffer();
            payload.Prepare(str);
            payload.PrepareApply();
            payload.QWrite(str);
            PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);


            CExcel exls = null;
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
                    Console.WriteLine("Excelfile <{0}> not found", workBookPath);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
               exls.CloseWorkBook();
            }

            ModuleResult result = new ModuleResult(ModuleId, request.variantId, request.kpiId, outputs);
            str = Ecodistrict.Messaging.Serialize.ToJsonString(result);
            payload = new TByteBuffer();
            payload.Prepare(str);
            payload.PrepareApply();
            payload.QWrite(str);
            PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);

            StartModuleResponse stmResp2 = new StartModuleResponse(ModuleId, request.variantId, request.kpiId, ModuleStatus.Success);
            str = Ecodistrict.Messaging.Serialize.ToJsonString(stmResp2);
            payload = new TByteBuffer();
            payload.Prepare(str);
            payload.PrepareApply();
            payload.QWrite(str);
            PublishedEvent.SignalEvent(TEventEntry.TEventKind.ekNormalEvent, payload.Buffer);

            return true;

        }

        //Private
    }
}

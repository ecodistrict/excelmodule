using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Yaml.Serialization;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace RenobuildModule
{
    class RenobuildModule : CExcelModule
    {
        public RenobuildModule()
        {
            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "";

            //List of kpis the module can calculate
            this.KpiList = new List<string> { "kpi1", "kpi2", "kpi3" };

            //Error handler
            this.ErrorRaised += RenobuildModule_ErrorRaised;
        }

        void RenobuildModule_ErrorRaised(object sender, ErrorMessage e)
        {
            Console.WriteLine(e.Message);
        }

        void RenobuildModule_ErrorRaised(object sender, Exception ex)
        {
            ErrorMessage em = new ErrorMessage();
            em.Message = ex.Message;
            RenobuildModule_ErrorRaised(sender, em);
        }

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            InputSpecification iSpec = new InputSpecification();

            switch(kpiId)
            {
                case "kpi1":
                    //Your input spec def.
                    break;
                case "kp12":
                    //Your input spec def.
                    break;
                default:
                    throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.", kpiId));
            }

            return iSpec;
        }

        protected override Outputs CalculateKpi(Dictionary<string, object> indata, string kpiId, CExcel exls) 
        {
            Outputs outputs = new Outputs();

            switch(kpiId)
            {
                case "kpi1":
                    //Do your calculations here.
                    break;
                case "kpi2":
                    //Do your calculations here.
                    break;
                default:
                    throw new ApplicationException(String.Format("No calcualtion procedure could be found for '{0}'", kpiId));
            }
            
            //tmp
            outputs.Add(new Kpi(1, "info", "unit"));

            return outputs;
        }

        public bool Init(string IMB_config_path, string Module_config_path)
        {
            try
            {
                Init_IMB(IMB_config_path);
                Init_Module(Module_config_path);
                return true;
            }
            catch(Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
                return false;
            }            
        }

        private void Init_IMB(string IMB_config_path)
        {
            try
            {
                var serializer = new YamlSerializer();
                var imb_settings = serializer.DeserializeFromFile(IMB_config_path, typeof(IMB_Settings))[0];

                this.ServerAdress = ((IMB_Settings)imb_settings).serverAdress;
                this.Port = ((IMB_Settings)imb_settings).port;
                this.Federation = ((IMB_Settings)imb_settings).federation;
                this.SubScribedEventName = ((IMB_Settings)imb_settings).subScribedEventName;
                this.PublishedEventName = ((IMB_Settings)imb_settings).publishedEventName;
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading the IMB configuration file", ex);
            }
        }

        private void Init_Module(string Module_config_path)
        {
            try
            {
                var serializer = new YamlSerializer();
                var module_settings = serializer.DeserializeFromFile(Module_config_path, typeof(Module_Settings))[0];

                this.ModuleName = ((Module_Settings)module_settings).name;
                this.Description = ((Module_Settings)module_settings).description;
                this.ModuleId = ((Module_Settings)module_settings).moduleId;
                this.workBookPath = ((Module_Settings)module_settings).path;
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading the module configuration file", ex);
            }
        }
    }
}

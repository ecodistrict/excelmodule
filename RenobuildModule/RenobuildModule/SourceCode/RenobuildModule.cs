using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Yaml.Serialization;
using Ecodistrict.Messaging;

namespace RenobuildModule
{
    class RenobuildModule //: ExcelModule
    {
        protected InputSpecification GetInputSpec(string kpiId)
        {
            InputSpecification iSpec = new InputSpecification();

            switch(kpiId)
            {
                case "something":
                    //Your input spec def.
                    break;
                case "somethingElse":
                    //Your input spec def.
                    break;
                default:
                    throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.", kpiId));
            }

            return iSpec;
        }

        protected Outputs CalculateKPI(Dictionary<string,object> indata, string kpiId)  //CExcel excelDoc
        {
            Outputs outputs = new Outputs();

            switch(kpiId)
            {
                case "something":
                    //Do your calculations here.
                    break;
                case "somethingElse":
                    //Do your calculations here.
                    break;
                default:
                    throw new ApplicationException(String.Format("No calcualtion procedure could be found for '{0}'", kpiId));
            }
            
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
                Console.WriteLine(ex.Message);
                return false;
            }            
        }

        public void Init_IMB(string IMB_config_path)
        {
            try
            {
                var serializer = new YamlSerializer();
                var imb_settings = serializer.DeserializeFromFile(IMB_config_path, typeof(IMB_Settings))[0];
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading the IMB configuration file", ex);
            }
        }

        public void Init_Module(string Module_config_path)
        {
            try
            {
                var serializer = new YamlSerializer();
                var module_settings = serializer.DeserializeFromFile(Module_config_path, typeof(Module_Settings))[0];
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading the module configuration file", ex);
            }
        }
    }
}

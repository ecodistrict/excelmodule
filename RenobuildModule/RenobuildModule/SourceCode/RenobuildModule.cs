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
        const string kpi_gwp = "change-of-global-warming-potential";
        const string kpi_peu = "change-of-primary-energy-use";

        public RenobuildModule()
        {
            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "";

            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_gwp, kpi_peu};

            //Error handler
            this.ErrorRaised += RenobuildModule_ErrorRaised;

            //Notification
            this.StatusMessage += RenobuildModule_StatusMessage;
        }

        void RenobuildModule_StatusMessage(object sender, StatusEventArg e)
        {
            Console.WriteLine(String.Format("Status message:\n\t{0}",e.StatusMessage));
        }

        void RenobuildModule_ErrorRaised(object sender, ErrorMessageEventArg e)
        {
            Console.WriteLine(String.Format("Error message: {0}", e.Message));
            if (e.SourceFunction != null & e.SourceFunction != "")
                Console.WriteLine(String.Format("\tIn source function: {0}", e.SourceFunction));
        }

        void RenobuildModule_ErrorRaised(object sender, Exception ex)
        {
            ErrorMessageEventArg em = new ErrorMessageEventArg();
            em.Message = ex.Message;
            RenobuildModule_ErrorRaised(sender, em);
        }

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            Options heat_sources = new Options();
            heat_sources.Add(new Option(value: "geothermal_heat_pump", label: "Geothermal heat pump"));
            heat_sources.Add(new Option(value: "district_heating", label: "District heating"));
            heat_sources.Add(new Option(value: "pellet_boiler", label: "Pellet boiler"));
            heat_sources.Add(new Option(value: "oil_boiler", label: "Oil boiler"));
            heat_sources.Add(new Option(value: "electric_boiler", label: "Electric boiler"));
            heat_sources.Add(new Option(value: "direct_electricity", label: "Direct electricity"));

            InputSpecification iSpec = new InputSpecification();

            //Vilka parametrar behövs för de olka kpi:erna?
            //Standardvärden

            switch(kpiId)
            {
                case kpi_gwp:
                case kpi_peu:
                    // - ## Common Properties
                    InputGroup commonProp = new InputGroup(label: "Common properties", order: 1);
                    commonProp.Add("period", new Number(label: "LCA calculation period", min: 1, unit: "years", order: 1));
                    //Applicable to disctrict heating system
                    InputGroup igAtDHS = new InputGroup(label: "Applicable to disctrict heating system (inputs required if district heating is used (before/after renovation))", order: 2);
                    igAtDHS.Add(key: "gwp_district", item: new Number(label: "Global warming potential of district heating", min: 0, unit: "g CO2 eq/kWh", order: 1));
                    igAtDHS.Add(key: "peu_district", item: new Number(label: "Primary energy use of district heating", min: 0, order: 2));
                    commonProp.Add("applicable_to_disctrict_heating_system", igAtDHS);

                    // - ## Building Specific
                    InputGroup buildingSpecificProp = new InputGroup(label: "Buildning specific properties", order: 1);
                    GeoJson buildning_specific_data = new GeoJson(label: "Geograpich data of buildnings", order: -1);
                    //Group: Applicable to building                 
                    InputGroup igAtB = new InputGroup(label: "Applicable to building", order: 1);
                    igAtB.Add(key: "period", item: new Number(label: "LCA calculation period", min: 1, unit: "years", order: 1));                   
                    igAtB.Add(key: "nr_apartments", item: new Number(label: "Number of apartments", min: 1, order: 2));
                    igAtB.Add(key: "heat_source_before", item: new Select(label: "Heat source before renovation", options: heat_sources, order: 3));
                    igAtB.Add(key: "heat_source_after", item: new Select(label: "Heat source after renovation", options: heat_sources, order: 4));
                    buildning_specific_data.Add("applicable_to_building", igAtB);
                    //Group: Applicable to building heating system
                    InputGroup igAtBHS = new InputGroup(label: "Applicable to building heating system (inputs required if heat source is replaced)", order: 3);
                    igAtBHS.Add(key: "ahd_after_renovation", item: new Number(label: "Annual heat demand after renovation (Required if heating system is replaced)", min: 0, unit: "kWh/year", order: 1));
                    igAtBHS.Add(key: "life_of_product", item: new Number(label: "Life of product (Practical time of life of the products and materials used)", min: 0, unit: "years", order: 2));
                    igAtBHS.Add(key: "design_capacity", item: new Number(label: "Design capacity (Required for pellets boiler and oil boiler)", min: 0, unit: "kW", order: 3));
                    igAtBHS.Add(key: "weight_of_bhd", item: new Number(label: "Weight of boiler/heat pump/district heating substation (Required except for direct electricity heating)", min: 0, unit: "kg", order: 4));
                    igAtBHS.Add(key: "depth_of_borehole", item: new Number(label: "Depth of borehole (For geothermal heat pump)", min: 0, unit: "m", order: 5));
                    igAtBHS.Add(key: "transport_to_building_truck", item: new Number(label: "Transport to building by truck (Distance from production site to building)", min: 0, unit: "km", order: 6));
                    igAtBHS.Add(key: "transport_to_building_train", item: new Number(label: "Transport to building by train (Distance from production site to building)", min: 0, unit: "km", order: 7));
                    igAtBHS.Add(key: "transport_to_building_ferry", item: new Number(label: "Transport to building by ferry (Distance from production site to building)", min: 0, unit: "km", order: 8));
                    buildning_specific_data.Add("applicable_to_building_heating_system", igAtBHS);
                    //Group: Applicable to pump in building heating system
                    InputGroup igAtPiBHS = new InputGroup(label: "Applicable to pump in building heating system (inputs required if circulation pump in building heating system is replaced)", order: 4);
                    igAtPiBHS.Add(key: "life_of_product", item: new Number(label: "Practical time of life of the products and materials used", min: 0, unit: "years", order: 1));
                    igAtPiBHS.Add(key: "design_pressure_head", item: new Number(label: "Design pressure head ()", min: 0, unit: "kPa", order: 2));
                    igAtPiBHS.Add(key: "design_flow_rate", item: new Number(label: "Design flow rate ()", min: 0, unit: "m\u00b3/h", order: 3));
                    igAtPiBHS.Add(key: "type_of_fcontrol_in_heating_system", item: new Number(label: "Type of flow control in heating system ()", min: 0, order: 4));
                    //igAtPiBHS.Add(key: "", item: new Number(label: "Weight ()", min: 0, unit: "", order: 5));  //In what form?
                    igAtPiBHS.Add(key: "transport_to_building_truck", item: new Number(label: "Transport to building by truck (Distance from production site to building)", min: 0, unit: "km", order: 6));
                    igAtPiBHS.Add(key: "transport_to_building_train", item: new Number(label: "Transport to building by train (Distance from production site to building)", min: 0, unit: "km", order: 7));
                    igAtPiBHS.Add(key: "transport_to_building_ferry", item: new Number(label: "Transport to building by ferry (Distance from production site to building)", min: 0, unit: "km", order: 8));
                    buildning_specific_data.Add("applicable_to_pump_in_building_heating_system", igAtPiBHS);
                    //...

                    buildingSpecificProp.Add(key: "buildnings", item: buildning_specific_data);


                    iSpec.Add("commonProp", buildingSpecificProp);
                    iSpec.Add("buildingSpecificProp", commonProp);

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
                case kpi_gwp:
                    //Do your calculations here.
                    break;
                case kpi_peu:
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

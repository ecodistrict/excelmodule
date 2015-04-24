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
            this.KpiList = new List<string> { kpi_gwp, kpi_peu };

            //Error handler
            this.ErrorRaised += RenobuildModule_ErrorRaised;

            //Notification
            this.StatusMessage += RenobuildModule_StatusMessage;
        }

        void RenobuildModule_StatusMessage(object sender, StatusEventArg e)
        {
            Console.WriteLine(String.Format("Status message:\n\t{0}", e.StatusMessage));
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


        List BuildingSpecificSpec(Options heat_sources)
        {
            // - ## Building Specific
            List buildning_specific_data = new List(label: "Buildings", order: 2);

            //Group: Applicable to building    
            buildning_specific_data.Add("applicable_to_building", ApplicableToBuilding(heat_sources));

            //Group: Applicable to district heating system
            buildning_specific_data.Add("applicable_to_disctrict_heating_system", ApplicableToDistrictHeatingSystem());

            //Group: Applicable to building heating system
            buildning_specific_data.Add("applicable_to_building_heating_system", ApplicableToBuildingHeatingSystem());

            //Group: Applicable to pump in building heating system
            buildning_specific_data.Add("applicable_to_pump_in_building_heating_system", ApplicableToPumpInBuildingHeatingSystem());
            //...

            return buildning_specific_data;
        }

        InputSpecification tmpSpec()
        {
            //For the selected kpi, create a input specification describing what data 
            //the module need in order to calculate the selected kpi.
            InputSpecification iSpec = new InputSpecification();
            //In this case the module needs 2 things.

            //A user specified age
            Number numberAge = new Number(
                label: "Age",
                min: 0,
                unit: "years");

            Options opt = new Options();
            opt.Add(new Option(value: "alp-cheese", label: "Alpk\u00e4se")); //Note the web-friendly string
            opt.Add(new Option(value: "edam-cheese", label: "Edammer"));
            Option brie = new Option(value: "brie-cheese", label: "Brie");
            opt.Add(brie);

            //And one of the above options of cheese-types. 
            //(The preselected value, "brie-cheese", is optional)
            Select selectCheseType = new Select(
                label: "Cheese type",
                options: opt,
                value: brie);

            //Add these components to the input specification.
            //(Note the choosed keys, its the keys that will be attached to the
            //data when the dashboard returns with the user specified data in
            //a StartModuleRequest.)
            iSpec.Add(
                key: "age",
                value: numberAge);

            iSpec.Add(
                key: "cheese-type",
                value: selectCheseType);

            return iSpec;

        }

        InputSpecification tmpSpec2()
        {
            InputSpecification iSpec = new InputSpecification();

            //Applicable to district heating system
            List aList = new List(label: "Applicable to district heating system (inputs required if district heating is used (before/after renovation))", order: 2);
            aList.Add(key: "gwp_district", item: new Number(label: "Global warming potential of district heating", min: 0, unit: "g CO2 eq/kWh", order: 1));
            aList.Add(key: "peu_district", item: new Number(label: "Primary energy use of district heating", min: 0, order: 2));

            iSpec.Add("aList", aList);

            return iSpec;

        }

        InputSpecification tmpSpec3()
        {
            InputSpecification iSpec = new InputSpecification();

            iSpec.Add("gwp_district", new Number(label: "Global warming potential of district heating", min: 0, unit: "g CO2 eq/kWh", order: 1));
            iSpec.Add("peu_district", new Number(label: "Primary energy use of district heating", min: 0, order: 2));

            return iSpec;

        }

        //..

        InputGroup HeatingSystem()
        {
            int order = 0;
            InputGroup igHeatingSystem = new InputGroup("Heating system");

            // Change Heating System
            string change_heating_system = "change_heating_system";
            string change_heating_system_lbl = "Replace building heating system";
            string ahd_after_renovation = "ahd_after_renovation";
            string ahd_after_renovation_lbl = "Annual heat demand after renovation";
            string heating_system_life_of_product = "heating_system_life_of_product";
            string heating_system_life_of_product_lbl = "Life of product (Practical time of life of the products and materials used)";
            string design_capacity = "design_capacity";
            string design_capacity_lbl = "Design capacity (Required for pellets boiler and oil boiler)";
            string weight_of_bhd = "weight_of_bhd";
            string weight_of_bhd_lbl = "Weight of boiler/heat pump/district heating substation (Required except for direct electricity heating)";
            string depth_of_borehole = "depth_of_borehole";
            string depth_of_borehole_lbl = "Depth of bore hole (For geothermal heat pump)";
            string heating_system_transport_to_building_truck = "heating_system_transport_to_building_truck";
            string heating_system_transport_to_building_truck_lbl = "Transport to building by truck (Distance from production site to building)";
            string heating_system_transport_to_building_train = "heating_system_transport_to_building_train";
            string heating_system_transport_to_building_train_lbl = "Transport to building by train (Distance from production site to building)";
            string heating_system_transport_to_building_ferry = "heating_system_transport_to_building_ferry";
            string heating_system_transport_to_building_ferry_lbl = "Transport to building by ferry (Distance from production site to building)";
            igHeatingSystem.Add(key: change_heating_system, item: new Checkbox(label: change_heating_system_lbl, order: ++order));
            igHeatingSystem.Add(key: ahd_after_renovation, item: new Number(label: ahd_after_renovation_lbl, min: 0, unit: "kWh/year", order: ++order));
            igHeatingSystem.Add(key: heating_system_life_of_product, item: new Number(label: heating_system_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igHeatingSystem.Add(key: design_capacity, item: new Number(label: design_capacity_lbl, min: 0, unit: "kW", order: ++order));
            igHeatingSystem.Add(key: weight_of_bhd, item: new Number(label: weight_of_bhd_lbl, min: 0, unit: "kg", order: ++order));
            igHeatingSystem.Add(key: depth_of_borehole, item: new Number(label: depth_of_borehole_lbl, min: 0, unit: "m", order: ++order));
            igHeatingSystem.Add(key: heating_system_transport_to_building_truck, item: new Number(label: heating_system_transport_to_building_truck_lbl, min: 0, unit: "km", order: ++order));
            igHeatingSystem.Add(key: heating_system_transport_to_building_train, item: new Number(label: heating_system_transport_to_building_train_lbl, min: 0, unit: "km", order: ++order));
            igHeatingSystem.Add(key: heating_system_transport_to_building_ferry, item: new Number(label: heating_system_transport_to_building_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Change Circulation Pump
            string change_circulationpump_in_heating_system = "change_circulationpump_in_heating_system";
            string change_circulationpump_in_heating_system_lbl = "Replace circulation pump in building heating system";
            string circulationpump_life_of_product = "circulationpump_life_of_product";
            string circulationpump_life_of_product_lbl = "Practical time of life of the products and materials used";
            string design_pressure_head = "design_pressure_head";
            string design_pressure_head_lbl = "Design pressure head ()";
            string design_flow_rate = "design_flow_rate";
            string design_flow_rate_lbl = "Design flow rate ()";
            string type_of_control_in_heating_system = "type_of_control_in_heating_system";
            string type_of_control_in_heating_system_lbl = "Type of flow control in heating system ()";
            string weight = "weight";
            string weight_lbl = "Weight ()";
            string circulationpump_transport_to_building_truck = "circulationpump_transport_to_building_truck";
            string circulationpump_transport_to_building_truck_lbl = "Transport to building by truck (Distance from production site to building)";
            string circulationpump_transport_to_building_train = "circulationpump_transport_to_building_train";
            string circulationpump_transport_to_building_train_lbl = "Transport to building by train (Distance from production site to building)";
            string circulationpump_transport_to_building_ferry = "circulationpump_transport_to_building_ferry";
            string circulationpump_transport_to_building_ferry_lbl = "Transport to building by ferry (Distance from production site to building)";
            igHeatingSystem.Add(key: change_circulationpump_in_heating_system, item: new Checkbox(label: change_circulationpump_in_heating_system_lbl, order: ++order));
            igHeatingSystem.Add(key: circulationpump_life_of_product, item: new Number(label: circulationpump_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igHeatingSystem.Add(key: design_pressure_head, item: new Number(label: design_pressure_head_lbl, min: 0, unit: "kPa", order: ++order));
            igHeatingSystem.Add(key: design_flow_rate, item: new Number(label: design_flow_rate_lbl, min: 0, unit: "m\u00b3/h", order: ++order));
            Options type_of_control_in_heating_system_opts = new Options();
            type_of_control_in_heating_system_opts.Add(new Option("constant", "Constant"));
            type_of_control_in_heating_system_opts.Add(new Option("variable", "Variable"));
            igHeatingSystem.Add(key: type_of_control_in_heating_system, item: new Select(label: type_of_control_in_heating_system_lbl, options: type_of_control_in_heating_system_opts, order: ++order));
            igHeatingSystem.Add(key: weight, item: new Number(label: weight_lbl, min: 0, unit: "kg", order: ++order));
            igHeatingSystem.Add(key: circulationpump_transport_to_building_truck, item: new Number(label: circulationpump_transport_to_building_truck_lbl, min: 0, unit: "km", order: ++order));
            igHeatingSystem.Add(key: circulationpump_transport_to_building_train, item: new Number(label: circulationpump_transport_to_building_train_lbl, min: 0, unit: "km", order: ++order));
            igHeatingSystem.Add(key: circulationpump_transport_to_building_ferry, item: new Number(label: circulationpump_transport_to_building_ferry_lbl, min: 0, unit: "km", order: ++order));
            //...

            return igHeatingSystem;
        }

        GeoJson BuildingSpecificSpecGeoJson2(Options heat_sources)
        {
            // - ## Building Specific
            GeoJson buildning_specific_data = new GeoJson(label: "Geographic data of buildings", order: 2);

            // Inputs required in all cases
            int order = 0;
            string heated_area = "heated_area";
            string heated_area_lbl = "Heated Area";
            string nr_apartments = "nr_apartments";
            string nr_apartments_lbl = "Number of apartments";
            string heat_source_before = "heat_source_before";
            string heat_source_before_lbl = "Heat source before renovation";
            string heat_source_after = "heat_source_after";
            string heat_source_after_lbl = "Heat source after renovation";
            buildning_specific_data.Add(key: heated_area, item: new Number(label: heated_area_lbl, min: 1, unit: "m\u00b2", order: ++order));
            buildning_specific_data.Add(key: nr_apartments, item: new Number(label: nr_apartments_lbl, min: 1, order: ++order));
            buildning_specific_data.Add(key: heat_source_before, item: new Select(label: heat_source_before_lbl, options: heat_sources, order: ++order));
            buildning_specific_data.Add(key: heat_source_after, item: new Select(label: heat_source_after_lbl, options: heat_sources, order: ++order));
            // If district heating is used (before/after renovation)
            string gwp_district = "gwp_district";
            string gwp_district_lbl = "Global warming potential of district heating (If district heating is used before/after renovation)";
            string peu_district = "peu_district";
            string peu_district_lbl = "Primary energy use of district heating (If district heating is used before/after renovation)";
            buildning_specific_data.Add(key: gwp_district, item: new Number(label: gwp_district_lbl, min: 0, unit: "g CO2 eq/kWh", order: ++order));
            buildning_specific_data.Add(key: peu_district, item: new Number(label: peu_district_lbl, min: 0, order: ++order));

            // Heating System
            buildning_specific_data.Add(key: "heating_system", item: HeatingSystem());


            return buildning_specific_data;
        }

        InputSpecification GetInputSpecificationGeoJson2(Options heat_sources)
        {
            InputSpecification iSpec = new InputSpecification();

            // - ## Common Properties
            iSpec.Add("commonProperties", CommonSpec());

            // - ## Building Specific
            iSpec.Add("buildingProperties", new InputGroup(label: "Building specific properties", order: 2));
            iSpec.Add("buildings", BuildingSpecificSpecGeoJson2(heat_sources));

            return iSpec;
        }

        //..
        InputGroup ApplicableToBuilding(Options heat_sources)
        {
            //Group: Applicable to building                 
            InputGroup igAtB = new InputGroup(label: "Applicable to building\nBelow inputs required in all cases", order: 1);
            //igAtB.Add(key: "period", item: new Number(label: "LCA calculation period", min: 1, unit: "years", order: 1));  //For every building or common?
            igAtB.Add(key: "nr_apartments", item: new Number(label: "Number of apartments", min: 1, order: 2));
            igAtB.Add(key: "heat_source_before", item: new Select(label: "Heat source before renovation", options: heat_sources, order: 3));
            igAtB.Add(key: "heat_source_after", item: new Select(label: "Heat source after renovation", options: heat_sources, order: 4));

            // Renovation properties
            igAtB.Add(key: "heating_system", item: new InputGroup("Heating system"));
            igAtB.Add(key: "change_heating_system", item: new Checkbox("Replace building heating system"));
            igAtB.Add(key: "change_circulationpump_in_heating_system", item: new Checkbox("Replace circulation pump in building heating system"));

            igAtB.Add(key: "renovate_shell", item: new InputGroup("Shell"));
            igAtB.Add(key: "renovate_shell_insulation_material_1", item: new Checkbox("Replace circulation pump in building heating system"));


            igAtB.Add(key: "renovate_ventilation", item: new InputGroup("Ventilation"));

            igAtB.Add(key: "renovate_radiators_pipes_electric", item: new InputGroup("Radiators, pipes and electric"));

            return igAtB;
        }

        InputGroup ApplicableToDistrictHeatingSystem()
        {
            //Applicable to district heating system
            InputGroup igAtDHS = new InputGroup(label: "Applicable to district heating system (inputs required if district heating is used (before/after renovation))", order: 2);
            igAtDHS.Add(key: "gwp_district", item: new Number(label: "Global warming potential of district heating", min: 0, unit: "g CO2 eq/kWh", order: 1));
            igAtDHS.Add(key: "peu_district", item: new Number(label: "Primary energy use of district heating", min: 0, order: 2));

            return igAtDHS;

        }

        InputGroup ApplicableToBuildingHeatingSystem()
        {
            //Group: Applicable to building heating system
            InputGroup igAtBHS = new InputGroup(label: "Applicable to building heating system (inputs required if heat source is replaced)", order: 3);
            igAtBHS.Add(key: "ahd_after_renovation", item: new Number(label: "Annual heat demand after renovation (Required if heating system is replaced)", min: 0, unit: "kWh/year", order: 1));
            igAtBHS.Add(key: "life_of_product", item: new Number(label: "Life of product (Practical time of life of the products and materials used)", min: 0, unit: "years", order: 2));
            igAtBHS.Add(key: "design_capacity", item: new Number(label: "Design capacity (Required for pellets boiler and oil boiler)", min: 0, unit: "kW", order: 3));
            igAtBHS.Add(key: "weight_of_bhd", item: new Number(label: "Weight of boiler/heat pump/district heating substation (Required except for direct electricity heating)", min: 0, unit: "kg", order: 4));
            igAtBHS.Add(key: "depth_of_borehole", item: new Number(label: "Depth of bore hole (For geothermal heat pump)", min: 0, unit: "m", order: 5));
            igAtBHS.Add(key: "transport_to_building_truck", item: new Number(label: "Transport to building by truck (Distance from production site to building)", min: 0, unit: "km", order: 6));
            igAtBHS.Add(key: "transport_to_building_train", item: new Number(label: "Transport to building by train (Distance from production site to building)", min: 0, unit: "km", order: 7));
            igAtBHS.Add(key: "transport_to_building_ferry", item: new Number(label: "Transport to building by ferry (Distance from production site to building)", min: 0, unit: "km", order: 8));

            return igAtBHS;
        }

        InputGroup ApplicableToPumpInBuildingHeatingSystem()
        {
            //Group: Applicable to pump in building heating system
            InputGroup igAtPiBHS = new InputGroup(label: "Applicable to pump in building heating system (inputs required if circulation pump in building heating system is replaced)", order: 4);
            igAtPiBHS.Add(key: "life_of_product", item: new Number(label: "Practical time of life of the products and materials used", min: 0, unit: "years", order: 1));
            igAtPiBHS.Add(key: "design_pressure_head", item: new Number(label: "Design pressure head ()", min: 0, unit: "kPa", order: 2));
            igAtPiBHS.Add(key: "design_flow_rate", item: new Number(label: "Design flow rate ()", min: 0, unit: "m\u00b3/h", order: 3));
            igAtPiBHS.Add(key: "type_of_control_in_heating_system", item: new Number(label: "Type of flow control in heating system ()", min: 0, order: 4));
            //igAtPiBHS.Add(key: "", item: new Number(label: "Weight ()", min: 0, unit: "", order: 5));  //...In what form?
            igAtPiBHS.Add(key: "transport_to_building_truck", item: new Number(label: "Transport to building by truck (Distance from production site to building)", min: 0, unit: "km", order: 6));
            igAtPiBHS.Add(key: "transport_to_building_train", item: new Number(label: "Transport to building by train (Distance from production site to building)", min: 0, unit: "km", order: 7));
            igAtPiBHS.Add(key: "transport_to_building_ferry", item: new Number(label: "Transport to building by ferry (Distance from production site to building)", min: 0, unit: "km", order: 8));
            //...

            return igAtPiBHS;

        }

        InputGroup CommonSpec()
        {
            // - ## Common Properties
            InputGroup commonProp = new InputGroup(label: "Common properties", order: 1);
            commonProp.Add("period", new Number(label: "LCA calculation period", min: 1, unit: "years", order: 1));
            ////Applicable to district heating system
            //commonProp.Add("applicable_to_disctrict_heating_system", ApplicableToDistrictHeatingSystem());

            return commonProp;
        }

        GeoJson BuildingSpecificSpecGeoJson(Options heat_sources)
        {
            // - ## Building Specific
            GeoJson buildning_specific_data = new GeoJson(label: "Geographic data of buildings", order: 2);

            //Group: Applicable to building    
            buildning_specific_data.Add("applicable_to_building", ApplicableToBuilding(heat_sources));

            //Group: Applicable to district heating system
            buildning_specific_data.Add("applicable_to_disctrict_heating_system", ApplicableToDistrictHeatingSystem());

            //Group: Applicable to building heating system
            buildning_specific_data.Add("applicable_to_building_heating_system", ApplicableToBuildingHeatingSystem());

            //Group: Applicable to pump in building heating system
            buildning_specific_data.Add("applicable_to_pump_in_building_heating_system", ApplicableToPumpInBuildingHeatingSystem());
            //...

            return buildning_specific_data;
        }

        InputSpecification GetInputSpecificationGeoJson(Options heat_sources)
        {
            InputSpecification iSpec = new InputSpecification();

            // - ## Common Properties
            iSpec.Add("commonProperties", CommonSpec());

            // - ## Building Specific
            iSpec.Add("buildingProperties", new InputGroup(label: "Building specific properties", order: 2));
            iSpec.Add("buildings", BuildingSpecificSpecGeoJson(heat_sources));

            return iSpec;
        }


        //..

        InputSpecification GetInputSpecificationList(Options heat_sources) //TODO
        {
            InputSpecification iSpec = new InputSpecification();

            // - ## Common Properties
            iSpec.Add("commonProperties", CommonSpec());

            // - ## Building Specific
            iSpec.Add("buildings", BuildingSpecificSpec(heat_sources));

            return iSpec;
        }


        //...
        InputSpecification GetInputSpecificationGeoJsonPlain(Options heat_sources)
        {
            InputSpecification iSpec = new InputSpecification();

            // - ## Common Properties
            iSpec.Add("commonProperties", CommonSpec());

            // - ## Building Specific
            iSpec.Add("buildings", BuildingSpecificSpecGeoJsonPlain(heat_sources));

            return iSpec;
        }

        GeoJson BuildingSpecificSpecGeoJsonPlain(Options heat_sources)
        {
            // - ## Building Specific
            GeoJson buildning_specific_data = new GeoJson(label: "Geographic data of buildings", order: 2);

            //Group: Applicable to building 
            ApplicableToBuilding(heat_sources, ref buildning_specific_data, 3);
            //buildning_specific_data.Add("applicable_to_building", ;

            ////Group: Applicable to district heating system
            //buildning_specific_data.Add("applicable_to_disctrict_heating_system", ApplicableToDistrictHeatingSystem());

            ////Group: Applicable to building heating system
            //buildning_specific_data.Add("applicable_to_building_heating_system", ApplicableToBuildingHeatingSystem());

            ////Group: Applicable to pump in building heating system
            //buildning_specific_data.Add("applicable_to_pump_in_building_heating_system", ApplicableToPumpInBuildingHeatingSystem());
            //...

            return buildning_specific_data;
        }


        int ApplicableToBuilding(Options heat_sources, ref GeoJson buildning_specific_data, int orderLast)
        {
            //Group: Applicable to building       
            buildning_specific_data.Add("applicable_to_building", item: new InputGroup(label: "Applicable to building", order: ++orderLast));
            //buildning_specific_data.Add(key: "applicable_to_buildingTXT", item: new Checkbox(label: "", value: false, order: ++orderLast));
            buildning_specific_data.Add(key: "nr_apartments", item: new Number(label: "Number of apartments", min: 1, order: ++orderLast));
            buildning_specific_data.Add(key: "heat_source_before", item: new Select(label: "Heat source before renovation", options: heat_sources, order: ++orderLast));
            buildning_specific_data.Add(key: "heat_source_after", item: new Select(label: "Heat source after renovation", options: heat_sources, order: ++orderLast));
            return orderLast;
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

            switch (kpiId)
            {
                case kpi_gwp:
                case kpi_peu:


                    // - ## Building Specific
                    //GeoJson buildning_specific_data = new GeoJson(label: "Geographic data of buildings", order: 2);
                    //InputGroup ipg = new InputGroup(label: "Applicable to building", order: 1);
                    //ipg.Add(key: "nr_apartments", item: new Number(label: "Number of apartments", min: 1, order: 2));
                    //buildning_specific_data.Add("aaa", ipg);
                    //buildning_specific_data.Add(key: "heat_source_before", item: new Select(label: "Heat source before renovation", options: heat_sources, order: 3));
                    //buildning_specific_data.Add(key: "heat_source_after", item: new Select(label: "Heat source after renovation", options: heat_sources, order: 4));
                    //iSpec.Add("buildings", buildning_specific_data);

                    //iSpec = GetInputSpecificationGeoJsonPlain(heat_sources);
                    iSpec = GetInputSpecificationGeoJson2(heat_sources);

                    //iSpec = GetInputSpecificationList(heat_sources);

                    //iSpec = tmpSpec2();

                    break;
                default:
                    throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.", kpiId));
            }

            return iSpec;
        }

        protected override Outputs CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls)
        {
            Outputs outputs = new Outputs();

            switch (kpiId)
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
            catch (Exception ex)
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

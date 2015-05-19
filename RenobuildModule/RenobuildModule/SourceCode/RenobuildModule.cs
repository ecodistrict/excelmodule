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
        #region Defines
        // - Kpis
        const string kpi_gwp = "change-of-global-warming-potential";
        const string kpi_peu = "change-of-primary-energy-use";

        Dictionary<string, InputSpecification> inputSpecifications;

        Options heat_sources;
        Options type_of_flow_control_in_heating_system_opts;

        Options type_of_insulation;
        Options type_of_fascade_system;
        Options type_of_windows;
        Options type_of_doors;

        Options type_of_ventilation_ducts_material;
        Options type_of_airflow_assembly;

        Options type_of_radiators;

        void DefineHeatSources()
        {
            try
            {
                heat_sources = new Options();
                heat_sources.Add(new Option(value: "geothermal_heat_pump", label: "Geothermal heat pump"));
                heat_sources.Add(new Option(value: "district_heating", label: "District heating"));
                heat_sources.Add(new Option(value: "pellet_boiler", label: "Pellet boiler"));
                heat_sources.Add(new Option(value: "oil_boiler", label: "Oil boiler"));
                heat_sources.Add(new Option(value: "electric_boiler", label: "Electric boiler"));
                heat_sources.Add(new Option(value: "direct_electricity", label: "Direct electricity"));
            }
            catch(Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }
        void DefineTypeOfFlowControl()
        {
            try
            {
	            type_of_flow_control_in_heating_system_opts = new Options();
	            type_of_flow_control_in_heating_system_opts.Add(new Option("constant", "Constant"));
	            type_of_flow_control_in_heating_system_opts.Add(new Option("variable", "Variable"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }

        void DefineTypeOfIsulation()
        {
            try
            {
	            type_of_insulation = new Options();
	            type_of_insulation.Add(new Option(value: "cellulose_fiber", label: "Cellulose fiber"));
	            type_of_insulation.Add(new Option(value: "glass_wool", label: "Glass wool"));
	            type_of_insulation.Add(new Option(value: "rock_wool", label: "Rock wool"));
	            type_of_insulation.Add(new Option(value: "polystyrene_foam", label: "Polystyrene foam"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }
        void DefineTypeOfFascadeSystem()
        {
            try
            {
                type_of_fascade_system = new Options();
                type_of_fascade_system.Add(new Option(value: @"A\8-15mm\Non ventilated\EPS\200mm", label: @"A\8-15mm\Non ventilated\EPS\200mm"));
                type_of_fascade_system.Add(new Option(value: @"B\4-8mm\Ventilated\Rock wool\50mm", label: @"B\4-8mm\Ventilated\Rock wool\50mm"));
                type_of_fascade_system.Add(new Option(value: @"B\4-8mm\Ventilated\Rock wool\80mm", label: @"B\4-8mm\Ventilated\Rock wool\80mm"));
                type_of_fascade_system.Add(new Option(value: @"B\4-8mm\Ventilated\Rock wool\100mm", label: @"B\4-8mm\Ventilated\Rock wool\100mm"));
                type_of_fascade_system.Add(new Option(value: @"C\8-12mm\Non ventilated\EPS\50mm", label: @"C\8-12mm\Non ventilated\EPS\50mm"));
                type_of_fascade_system.Add(new Option(value: @"C\8-12mm\Non ventilated\EPS\80mm", label: @"C\8-12mm\Non ventilated\EPS\80mm"));
                type_of_fascade_system.Add(new Option(value: @"C\8-12mm\Non ventilated\EPS\100mm", label: @"C\8-12mm\Non ventilated\EPS\100mm"));
                type_of_fascade_system.Add(new Option(value: @"D\20-mm\Non ventilated\Rock wool\50mm", label: @"D\20-mm\Non ventilated\Rock wool\50mm"));
                type_of_fascade_system.Add(new Option(value: @"D\20-mm\Non ventilated\Rock wool\80mm", label: @"D\20-mm\Non ventilated\Rock wool\80mm"));
                type_of_fascade_system.Add(new Option(value: @"D\20-mm\Non ventilated\Rock wool\100mm", label: @"D\20-mm\Non ventilated\Rock wool\100mm"));
                type_of_fascade_system.Add(new Option(value: @"E\10-15mm\Non ventilated\Rock wool\50mm", label: @"E\10-15mm\Non ventilated\Rock wool\50mm"));
                type_of_fascade_system.Add(new Option(value: @"E\10-15mmNon ventilated\Rock wool\80mm", label: @"E\10-15mmNon ventilated\Rock wool\80mm"));
                type_of_fascade_system.Add(new Option(value: @"E\10-15mm\Non ventilated\Rock wool\100mm", label: @"E\10-15mm\Non ventilated\Rock wool\100mm"));
                type_of_fascade_system.Add(new Option(value: @"E\10-15mm\Non ventilated\Rock wool, PIR\50+150mm", label: @"E\10-15mm\Non ventilated\Rock wool, PIR\50+150mm"));
                type_of_fascade_system.Add(new Option(value: @"F\4-8mm\Ventilated\Rock wool\80mm", label: @"F\4-8mm\Ventilated\Rock wool\80mm"));
                type_of_fascade_system.Add(new Option(value: @"F\4-8mm\Ventilated\Rock wool\100mm", label: @"F\4-8mm\Ventilated\Rock wool\100mm"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }
        void DefineTypeOfWindows()
        {
            try
            {
                type_of_windows = new Options();
                type_of_windows.Add(new Option(value: "aluminium", label: "Aluminium"));
                type_of_windows.Add(new Option(value: "plastic", label: "Plastic"));
                type_of_windows.Add(new Option(value: "wood_metal", label: "Wood-metal"));
                type_of_windows.Add(new Option(value: "wood", label: "Wood"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }
        void DefineTypeOfDoors()
        {
            try
            {
                type_of_doors = new Options();
                type_of_doors.Add(new Option(value: "front_door_wood_aluminium", label: "Front door wood-aluminium"));
                type_of_doors.Add(new Option(value: "front_door_wood_glass", label: "Front door wood-glass"));
                type_of_doors.Add(new Option(value: "balcony_glass_wood", label: "Balcony glass-wood"));
                type_of_doors.Add(new Option(value: "balcony_glass_wood_aluminium", label: "Balcony glass-wood-aluminium"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }
        
        void DefineTypeOfVentilationDuctsMaterial()
        {
            try
            {
                type_of_ventilation_ducts_material = new Options();
                type_of_ventilation_ducts_material.Add(new Option(value: "steel", label: "Steel"));
                type_of_ventilation_ducts_material.Add(new Option(value: "polyethylene", label: "Polyethylene"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }
        void DefineTypeOfAirflowAssembly()
        {
            try
            {
                type_of_airflow_assembly = new Options();
                type_of_airflow_assembly.Add(new Option(value: "exhaust_air_unit", label: "Exhaust air unit"));
                type_of_airflow_assembly.Add(new Option(value: "ventilation_unit_with_heat_recovery", label: "Ventilation unit with heat recovery"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }

        void DefineTypeOfRadiators()
        {
            try
            {
                type_of_radiators = new Options();
                type_of_radiators.Add(new Option(value: "waterborne", label: "Waterborne"));
                type_of_radiators.Add(new Option(value: "direct_electricity", label: "Direct electricity"));
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }

        void DefineInputSpecifications()
        {
            try
            {
                inputSpecifications = new Dictionary<string, InputSpecification>();

                //GeoJson
                //inputSpecifications.Add(kpi_gwp, GetInputSpecificationGeoJson());
                //inputSpecifications.Add(kpi_peu, GetInputSpecificationGeoJson());

                //One Building (simple)
                inputSpecifications.Add(kpi_gwp, GetInputSpecificationOneBuilding());
                inputSpecifications.Add(kpi_peu, GetInputSpecificationOneBuilding());
            }
            catch (System.Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }

        #region Input Specification (Names and Labels)
        // - Input Specification (Names and Labels)
        #region Common Properties
        // Common
        string common_properties = "common_properties";
        string lca_calculation_period = "lca_calculation_period";

        #endregion

        #region Building specific Properties
        // Building specific Properties
        string buildings = "buildings";
        #region Building Common
        //Building Common
        // Inputs required in all cases
        string heated_area = "heated_area";
        string heated_area_lbl = "Heated area";
        string nr_apartments = "nr_apartments";
        string nr_apartments_lbl = "Number of apartments";
        string heat_source_before = "heat_source_before";
        string heat_source_before_lbl = "Heat source before renovation";
        string heat_source_after = "heat_source_after";
        string heat_source_after_lbl = "Heat source after renovation";
        // If district heating is used (before/after renovation)
        string gwp_district = "gwp_district";
        string gwp_district_lbl = "Global warming potential of district heating. Required if building uses district heating (before/after renovation). Impact per unit energy delivered to building, i.e. including distribution losses.";
        string peu_district = "peu_district";
        string peu_district_lbl = "Primary energy use of district heating (before/after renovation). Impact per unit energy delivered to building, i.e. including distribution losses.";
        #endregion

        #region Heating System
        //Heating System
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

        // Change Circulation Pump
        string change_circulationpump_in_heating_system = "change_circulationpump_in_heating_system";
        string change_circulationpump_in_heating_system_lbl = "Replace circulation pump in building heating system";
        string circulationpump_life_of_product = "circulationpump_life_of_product";
        string circulationpump_life_of_product_lbl = "Practical time of life of the products and materials used";
        string design_pressure_head = "design_pressure_head";
        string design_pressure_head_lbl = "Design pressure head";
        string design_flow_rate = "design_flow_rate";
        string design_flow_rate_lbl = "Design flow rate";
        string type_of_control_in_heating_system = "type_of_control_in_heating_system";
        string type_of_control_in_heating_system_lbl = "Type of flow control in heating system";
        string weight = "weight";
        string weight_lbl = "Weight of new pump";
        string circulationpump_transport_to_building_truck = "circulationpump_transport_to_building_truck";
        string circulationpump_transport_to_building_truck_lbl = "Transport to building by truck (Distance from production site to building)";
        string circulationpump_transport_to_building_train = "circulationpump_transport_to_building_train";
        string circulationpump_transport_to_building_train_lbl = "Transport to building by train (Distance from production site to building)";
        string circulationpump_transport_to_building_ferry = "circulationpump_transport_to_building_ferry";
        string circulationpump_transport_to_building_ferry_lbl = "Transport to building by ferry (Distance from production site to building)";
        #endregion

        #region Building Shell
        //Building Shell
        // Insulation material 1
        string change_insulation_material_1 = "change_insulation_material_1";
        string change_insulation_material_1_lbl = "Use insulation material 1";
        string insulation_material_1_life_of_product = "insulation_material_1_life_of_product";
        string insulation_material_1_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string insulation_material_1_type_of_insulation = "insulation_material_1_type_of_insulation";
        string insulation_material_1_type_of_insulation_lbl = "Type of insulation";
        string insulation_material_1_change_in_annual_heat_demand_due_to_insulation = "insulation_material_1_change_in_annual_heat_demand_due_to_insulation";
        string insulation_material_1_change_in_annual_heat_demand_due_to_insulation_lbl = "Change in annual heat demand due to insulation (an energy saving is given as a negative value)";
        string insulation_material_1_amount_of_new_insulation_material = "insulation_material_1_amount_of_new_insulation_material";
        string insulation_material_1_amount_of_new_insulation_material_lbl = "Amount of new insulation material (required if renovation includes new insulation material)";
        string insulation_material_1_transport_to_building_by_truck = "insulation_material_1_transport_to_building_by_truck";
        string insulation_material_1_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string insulation_material_1_transport_to_building_by_train = "insulation_material_1_transport_to_building_by_train";
        string insulation_material_1_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string insulation_material_1_transport_to_building_by_ferry = "insulation_material_1_transport_to_building_by_ferry";
        string insulation_material_1_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Insulation material 2
        string change_insulation_material_2 = "change_insulation_material_2";
        string change_insulation_material_2_lbl = "Use insulation material 2";
        string insulation_material_2_life_of_product = "insulation_material_2_life_of_product";
        string insulation_material_2_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string insulation_material_2_type_of_insulation = "insulation_material_2_type_of_insulation";
        string insulation_material_2_type_of_insulation_lbl = "Type of insulation";
        string insulation_material_2_change_in_annual_heat_demand_due_to_insulation = "insulation_material_2_change_in_annual_heat_demand_due_to_insulation";
        string insulation_material_2_change_in_annual_heat_demand_due_to_insulation_lbl = "Change in annual heat demand due to insulation (an energy saving is given as a negative value)";
        string insulation_material_2_amount_of_new_insulation_material = "insulation_material_2_amount_of_new_insulation_material";
        string insulation_material_2_amount_of_new_insulation_material_lbl = "Amount of new insulation material (required if renovation includes new insulation material)";
        string insulation_material_2_transport_to_building_by_truck = "insulation_material_2_transport_to_building_by_truck";
        string insulation_material_2_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string insulation_material_2_transport_to_building_by_train = "insulation_material_2_transport_to_building_by_train";
        string insulation_material_2_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string insulation_material_2_transport_to_building_by_ferry = "insulation_material_2_transport_to_building_by_ferry";
        string insulation_material_2_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Fascade system
        string change_fascade_system = "change_fascade_system";
        string change_fascade_system_lbl = "Change fascade";
        string fascade_system_life_of_product = "fascade_system_life_of_product";
        string fascade_system_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string fascade_system_type_fascade_system = "fascade_system_type_fascade_system";
        string fascade_system_type_of_fascade_system_lbl = "Type of facade system";
        string fascade_system_change_in_annual_heat_demand_due_to_fascade_system = "fascade_system_change_in_annual_heat_demand_due_to_fascade_system";
        string fascade_system_change_in_annual_heat_demand_due_to_fascade_system_lbl = "Change in annual heat demand due to fascade system (an energy saving is given as a negative value)";
        string fascade_system_area_of_new_fascade_system = "fascade_system_amount_of_new_insulation_material";
        string fascade_system_area_of_new_fascade_system_lbl = "Area of new facade system (required if renovation includes new facade system)";
        string fascade_system_transport_to_building_by_truck = "fascade_system_transport_to_building_by_truck";
        string fascade_system_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string fascade_system_transport_to_building_by_train = "fascade_system_transport_to_building_by_train";
        string fascade_system_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string fascade_system_transport_to_building_by_ferry = "fascade_system_transport_to_building_by_ferry";
        string fascade_system_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";


        // Windows
        string change_windows = "change_windows";
        string change_windows_lbl = "Change windows";
        string windows_life_of_product = "windows_life_of_product";
        string windows_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string windows_type_windows = "windows_type_windows";
        string windows_type_of_windows_lbl = "Material in frame";
        string windows_change_in_annual_heat_demand_due_to_windows = "windows_change_in_annual_heat_demand_due_to_windows";
        string windows_change_in_annual_heat_demand_due_to_windows_lbl = "Change in annual heat demand due to windows (an energy saving is given as a negative value)";
        string windows_area_of_new_windows = "windows_amount_of_new_insulation_material";
        string windows_area_of_new_windows_lbl = "Area of windows (required if renovation includes new windows)";
        string windows_transport_to_building_by_truck = "windows_transport_to_building_by_truck";
        string windows_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string windows_transport_to_building_by_train = "windows_transport_to_building_by_train";
        string windows_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string windows_transport_to_building_by_ferry = "windows_transport_to_building_by_ferry";
        string windows_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Doors
        string change_doors = "change_doors";
        string change_doors_lbl = "Change doors";
        string doors_life_of_product = "doors_life_of_product";
        string doors_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string doors_type_doors = "doors_type_doors";
        string doors_type_of_doors_lbl = "Type of doors";
        string doors_change_in_annual_heat_demand_due_to_doors = "doors_change_in_annual_heat_demand_due_to_insulation";
        string doors_change_in_annual_heat_demand_due_to_doors_lbl = "Change in annual heat demand due to doors (an energy saving is given as a negative value)";
        string doors_number_of_new_front_doors = "doors_amount_of_new_insulation_material";
        string doors_number_of_new_front_doors_lbl = "Number of new front doors (required if renovation includes new doors)";
        string doors_transport_to_building_by_truck = "doors_transport_to_building_by_truck";
        string doors_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string doors_transport_to_building_by_train = "doors_transport_to_building_by_train";
        string doors_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string doors_transport_to_building_by_ferry = "doors_transport_to_building_by_ferry";
        string doors_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";
        #endregion

        #region Ventilation
        // Ventilation
        // Ventilation ducts
        string change_ventilation_ducts = "change_ventilation_ducts";
        string change_ventilation_ducts_lbl = "Change ventilation ducts";
        string ventilation_ducts_life_of_product = "ventilation_ducts_life_of_product";
        string ventilation_ducts_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string ventilation_ducts_type_of_material = "ventilation_ducts_type_of_material";
        string ventilation_ducts_type_of_material_lbl = "Material in ventilation ducts";
        string ventilation_ducts_weight_of_ventilation_ducts = "ventilation_ducts_weight_of_ventilation_ducts (Required if renovation includes new ventilation ducts)";
        string ventilation_ducts_weight_of_ventilation_ducts_lbl = "Weight of ventilation ducts";
        string ventilation_ducts_transport_to_building_by_truck = "ventilation_ducts_transport_to_building_by_truck";
        string ventilation_ducts_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string ventilation_ducts_transport_to_building_by_train = "ventilation_ducts_transport_to_building_by_train";
        string ventilation_ducts_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string ventilation_ducts_transport_to_building_by_ferry = "ventilation_ducts_transport_to_building_by_ferry";
        string ventilation_ducts_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Airflow assembly
        string change_airflow_assembly = "change_airflow_assembly";
        string change_airflow_assembly_lbl = "Change airflow assembly";
        string airflow_assembly_life_of_product = "airflow_assembly_life_of_product";
        string airflow_assembly_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string airflow_assembly_type_of_airflow_assembly = "airflow_assembly_type_of_insulation";
        string airflow_assembly_type_of_airflow_assembly_lbl = "Type of airflow assembly";
        string airflow_assembly_design_airflow_exhaust_air = "airflow_assembly_design_airflow_exhaust_air";
        string airflow_assembly_design_airflow_exhaust_air_lbl = "Design airflow (exhaust air)";
        string airflow_assembly_transport_to_building_by_truck = "airflow_assembly_transport_to_building_by_truck";
        string airflow_assembly_transport_to_building_by_truck_lbl = "Transport of airflow assembly to building by truck (distance from production site to building)";
        string airflow_assembly_transport_to_building_by_train = "airflow_assembly_transport_to_building_by_train";
        string airflow_assembly_transport_to_building_by_train_lbl = "Transport of airflow assembly to building by train (distance from production site to building)";
        string airflow_assembly_transport_to_building_by_ferry = "airflow_assembly_transport_to_building_by_ferry";
        string airflow_assembly_transport_to_building_by_ferry_lbl = "Transport of airflow assembly to building by ferry (distance from production site to building)";

        // Air distribution housings and silencer
        string change_air_distribution_housings_and_silencers = "change_air_distribution_housings_and_silencers";
        string change_air_distribution_housings_and_silencers_lbl = "Change air distribution housings and silencers";
        string air_distribution_housings_and_silencers_life_of_product = "air_distribution_housings_and_silencers_life_of_product";
        string air_distribution_housings_and_silencers_life_of_product_lbl = "Life of air distribution housings and silencers (practical time of life of the products and materials used)";
        string air_distribution_housings_and_silencers_transport_to_building_by_truck = "air_distribution_housings_and_silencers_transport_to_building_by_truck";
        string air_distribution_housings_and_silencers_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string air_distribution_housings_and_silencers_transport_to_building_by_train = "air_distribution_housings_and_silencers_transport_to_building_by_train";
        string air_distribution_housings_and_silencers_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string air_distribution_housings_and_silencers_transport_to_building_by_ferry = "air_distribution_housings_and_silencers_transport_to_building_by_ferry";
        string air_distribution_housings_and_silencers_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";
        
        //Ventilation renovation
        string ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation = "ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation";
        string ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation_lbl = "Change in annual heat demand due ventilation systems renovation (an energy saving is given as a negative value)";
        string ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation = "ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation";
        string ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation_lbl = "Change in annual electricity demand due ventilation systems renovation (an energy saving is given as a negative value)";
       
        #endregion

        #region Radiators, pipes and electricity
        // Radiators, pipes and electricity
        // Radiators
        string change_radiators = "change_radiators";
        string change_radiators_lbl = "Change radiators";
        string radiators_life_of_product = "radiators_life_of_product";
        string radiators_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string radiators_type_of_radiators = "radiators_type_of_radiators";
        string radiators_type_of_radiators_lbl = "Type of radiators";
        string radiators_weight_of_radiators = "radiators_weight_of_radiators";
        string radiators_weight_of_radiators_lbl = "Weight of new radiators";
        string radiators_transport_to_building_by_truck = "radiators_transport_to_building_by_truck";
        string radiators_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string radiators_transport_to_building_by_train = "radiators_transport_to_building_by_train";
        string radiators_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string radiators_transport_to_building_by_ferry = "radiators_transport_to_building_by_ferry";
        string radiators_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Copper
        string change_piping_copper = "change_piping_copper";
        string change_piping_copper_lbl = "Change copper pipes";
        string piping_copper_life_of_product = "piping_copper_life_of_product";
        string piping_copper_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_copper_weight_of_copper_pipes = "piping_copper_weight_of_copper_pipes";
        string piping_copper_weight_of_copper_pipes_lbl = "Weight of new pipes";
        string piping_copper_transport_to_building_by_truck = "piping_copper_transport_to_building_by_truck";
        string piping_copper_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_copper_transport_to_building_by_train = "piping_copper_transport_to_building_by_train";
        string piping_copper_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_copper_transport_to_building_by_ferry = "piping_copper_transport_to_building_by_ferry";
        string piping_copper_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - PEX
        string change_piping_pex = "change_piping_pex";
        string change_piping_pex_lbl = "Change PEX pipes";
        string piping_pex_life_of_product = "piping_pex_life_of_product";
        string piping_pex_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_pex_weight_of_pex_pipes = "piping_pex_weight_of_pex_pipes";
        string piping_pex_weight_of_pex_pipes_lbl = "Weight of new pipes";
        string piping_pex_transport_to_building_by_truck = "piping_pex_transport_to_building_by_truck";
        string piping_pex_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_pex_transport_to_building_by_train = "piping_pex_transport_to_building_by_train";
        string piping_pex_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_pex_transport_to_building_by_ferry = "piping_pex_transport_to_building_by_ferry";
        string piping_pex_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - PP
        string change_piping_pp = "change_piping_pp";
        string change_piping_pp_lbl = "Change PP pipes";
        string piping_pp_life_of_product = "piping_pp_life_of_product";
        string piping_pp_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_pp_weight_of_pp_pipes = "piping_pp_weight_of_pp_pipes";
        string piping_pp_weight_of_pp_pipes_lbl = "Weight of new pipes";
        string piping_pp_transport_to_building_by_truck = "piping_pp_transport_to_building_by_truck";
        string piping_pp_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_pp_transport_to_building_by_train = "piping_pp_transport_to_building_by_train";
        string piping_pp_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_pp_transport_to_building_by_ferry = "piping_pp_transport_to_building_by_ferry";
        string piping_pp_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Cast Iron
        string change_piping_cast_iron = "change_piping_cast_iron";
        string change_piping_cast_iron_lbl = "Change cast iron pipes";
        string piping_cast_iron_life_of_product = "piping_cast_iron_life_of_product";
        string piping_cast_iron_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_cast_iron_weight_of_cast_iron_pipes = "piping_cast_iron_weight_of_cast_iron_pipes";
        string piping_cast_iron_weight_of_cast_iron_pipes_lbl = "Weight of new pipes";
        string piping_cast_iron_transport_to_building_by_truck = "piping_cast_iron_transport_to_building_by_truck";
        string piping_cast_iron_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_cast_iron_transport_to_building_by_train = "piping_cast_iron_transport_to_building_by_train";
        string piping_cast_iron_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_cast_iron_transport_to_building_by_ferry = "piping_cast_iron_transport_to_building_by_ferry";
        string piping_cast_iron_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Galvanized Steel
        string change_piping_galvanized_steel = "change_piping_galvanized_steel";
        string change_piping_galvanized_steel_lbl = "Change galvanized steel pipes";
        string piping_galvanized_steel_life_of_product = "piping_galvanized_steel_life_of_product";
        string piping_galvanized_steel_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_galvanized_steel_weight_of_galvanized_steel_pipes = "piping_galvanized_steel_weight_of_galvanized_steel_pipes";
        string piping_galvanized_steel_weight_of_galvanized_steel_pipes_lbl = "Weight of new pipes";
        string piping_galvanized_steel_transport_to_building_by_truck = "piping_galvanized_steel_transport_to_building_by_truck";
        string piping_galvanized_steel_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_galvanized_steel_transport_to_building_by_train = "piping_galvanized_steel_transport_to_building_by_train";
        string piping_galvanized_steel_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_galvanized_steel_transport_to_building_by_ferry = "piping_galvanized_steel_transport_to_building_by_ferry";
        string piping_galvanized_steel_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Relining
        string change_piping_relining = "change_piping_relining";
        string change_piping_relining_lbl = "Relining of pipes";
        string piping_relining_life_of_product = "piping_relining_life_of_product";
        string piping_relining_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_relining_weight_of_relining_pipes = "piping_relining_weight_of_relining_pipes";
        string piping_relining_weight_of_relining_pipes_lbl = "Weight of new pipes";
        string piping_relining_transport_to_building_by_truck = "piping_relining_transport_to_building_by_truck";
        string piping_relining_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_relining_transport_to_building_by_train = "piping_relining_transport_to_building_by_train";
        string piping_relining_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_relining_transport_to_building_by_ferry = "piping_relining_transport_to_building_by_ferry";
        string piping_relining_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Electrical wiring
        string change_electrical_wiring = "change_electrical_wiring";
        string change_electrical_wiring_lbl = "Replace electrical wiring";
        string electrical_wiring_life_of_product = "electrical_wiring_life_of_product";
        string electrical_wiring_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string electrical_wiring_weight_of_electrical_wiring = "electrical_wiring_weight_of_electrical_wiring";
        string electrical_wiring_weight_of_electrical_wiring_lbl = "Weight of new wires";
        string electrical_wiring_transport_to_building_by_truck = "electrical_wiring_transport_to_building_by_truck";
        string electrical_wiring_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string electrical_wiring_transport_to_building_by_train = "electrical_wiring_transport_to_building_by_train";
        string electrical_wiring_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string electrical_wiring_transport_to_building_by_ferry = "electrical_wiring_transport_to_building_by_ferry";
        string electrical_wiring_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        #endregion

        #endregion

        #endregion

        #endregion

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

            //Define parameter options
            DefineHeatSources();
            DefineTypeOfFlowControl();

            DefineTypeOfIsulation();
            DefineTypeOfFascadeSystem();
            DefineTypeOfWindows();
            DefineTypeOfDoors();

            DefineTypeOfVentilationDuctsMaterial();
            DefineTypeOfAirflowAssembly();

            DefineTypeOfRadiators();

            //Define the input specification for the different kpis
            DefineInputSpecifications();
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

        InputSpecification GetInputSpecificationGeoJson()
        {
            InputSpecification iSpec = new InputSpecification();

            // - ## Common Properties
            iSpec.Add(common_properties, CommonSpec());

            // - ## Building Specific
            string description = "Building specific properties (Use the geojson-upload functionality below the map, in order to upload you buildings)";

            iSpec.Add("buildingProperties", new InputGroup(label: description, order: 2));
            iSpec.Add(buildings, BuildingSpecificSpecGeoJson());

            return iSpec;
        }

        InputGroup CommonSpec()
        {
            // - ## Common Properties
            InputGroup commonProp = new InputGroup(label: "Common properties", order: 1);
            commonProp.Add(lca_calculation_period, new Number(label: "LCA calculation period", min: 1, unit: "years", order: 1));
            ////Applicable to district heating system
            //commonProp.Add("applicable_to_disctrict_heating_system", ApplicableToDistrictHeatingSystem());

            return commonProp;
        }

        GeoJson BuildingSpecificSpecGeoJson()
        {
            // - ## Building Specific
            GeoJson buildning_specific_data = new GeoJson(label: "Geographic data of buildings", order: 2);

            int order = 0;

            // Instructions
            string intstr = "";
            intstr = "Fill in the building specific data below. "; 
            intstr += "Use the checkboxes to indicate what types of renovation procedures you want to perform for this alternative. ";
            intstr += "You need to fill in the  building properties as well as the parameters under checked checkboxes. ";
            intstr += "If this is the as-is step leave all checkboxes unchecked and fill in only the building properties. ";
            intstr += "If multiple buildings have common properties you may select all of them and assign them values simultaniously. ";
            InputGroup instructions = new InputGroup(label: intstr, order: ++order);
            buildning_specific_data.Add(key: "instructions", item: instructions);

            // Building Common
            buildning_specific_data.Add(key: "properties", item: BuildingProperties(++order));

            // Heating System
            buildning_specific_data.Add(key: "heating_system", item: HeatingSystem(++order));

            // Building Shell
            buildning_specific_data.Add(key: "building_shell", item: BuildingShell(++order));
            
            // Ventilation System
            buildning_specific_data.Add(key: "ventilation_system", item: VentilationSystem(++order));

            // Radiators, pipes and electricity
            buildning_specific_data.Add(key: "radiators_pipes_and_electricity", item: RadiatorsPipesElectricity(++order));

            return buildning_specific_data;
        }

        InputGroup BuildingProperties(int ipgOrder = -1)
        {
            int order = 0;
            InputGroup igBuildingCommon = new InputGroup("Building Properties", order: ++ipgOrder);

            // Inputs required in all cases
            igBuildingCommon.Add(key: heated_area, item: new Number(label: heated_area_lbl, min: 1, unit: "m\u00b2", order: ++order));
            igBuildingCommon.Add(key: nr_apartments, item: new Number(label: nr_apartments_lbl, min: 1, order: ++order));
            igBuildingCommon.Add(key: heat_source_before, item: new Select(label: heat_source_before_lbl, options: heat_sources, order: ++order));
            igBuildingCommon.Add(key: heat_source_after, item: new Select(label: heat_source_after_lbl, options: heat_sources, order: ++order));

            // If district heating is used (before/after renovation)
            igBuildingCommon.Add(key: gwp_district, item: new Number(label: gwp_district_lbl, min: 0, unit: "g CO2 eq/kWh", order: ++order));
            igBuildingCommon.Add(key: peu_district, item: new Number(label: peu_district_lbl, min: 0, unit: "kWh/kWh", order: ++order));


            return igBuildingCommon;
        }

        InputGroup HeatingSystem(int ipgOrder = -1)
        {
            int order = 0;
            InputGroup igHeatingSystem = new InputGroup("Heating system", ++ipgOrder);

            // Change Heating System
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
            igHeatingSystem.Add(key: change_circulationpump_in_heating_system, item: new Checkbox(label: change_circulationpump_in_heating_system_lbl, order: ++order));
            igHeatingSystem.Add(key: circulationpump_life_of_product, item: new Number(label: circulationpump_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igHeatingSystem.Add(key: design_pressure_head, item: new Number(label: design_pressure_head_lbl, min: 0, unit: "kPa", order: ++order));
            igHeatingSystem.Add(key: design_flow_rate, item: new Number(label: design_flow_rate_lbl, min: 0, unit: "m\u00b3/h", order: ++order));            
            igHeatingSystem.Add(key: type_of_control_in_heating_system, item: new Select(label: type_of_control_in_heating_system_lbl, options: type_of_flow_control_in_heating_system_opts, order: ++order));
            igHeatingSystem.Add(key: weight, item: new Number(label: weight_lbl, min: 0, unit: "kg", order: ++order));
            igHeatingSystem.Add(key: circulationpump_transport_to_building_truck, item: new Number(label: circulationpump_transport_to_building_truck_lbl, min: 0, unit: "km", order: ++order));
            igHeatingSystem.Add(key: circulationpump_transport_to_building_train, item: new Number(label: circulationpump_transport_to_building_train_lbl, min: 0, unit: "km", order: ++order));
            igHeatingSystem.Add(key: circulationpump_transport_to_building_ferry, item: new Number(label: circulationpump_transport_to_building_ferry_lbl, min: 0, unit: "km", order: ++order));
            
            return igHeatingSystem;
        }

        InputGroup BuildingShell(int ipgOrder = -1)
        {

            int order = 0;
            InputGroup igBuildingShell = new InputGroup("Building Shell",++ipgOrder);

            // Insulation material 1
            igBuildingShell.Add(key: change_insulation_material_1, item: new Checkbox(label: change_insulation_material_1_lbl, order: ++order));
            igBuildingShell.Add(key: insulation_material_1_life_of_product, item: new Number(label: insulation_material_1_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igBuildingShell.Add(key: insulation_material_1_type_of_insulation, item: new Select(label: insulation_material_1_type_of_insulation_lbl, options: type_of_insulation, order: ++order));
            igBuildingShell.Add(key: insulation_material_1_change_in_annual_heat_demand_due_to_insulation, item: new Number(label: insulation_material_1_change_in_annual_heat_demand_due_to_insulation_lbl, unit: "kWh/year", order: ++order));
            igBuildingShell.Add(key: insulation_material_1_amount_of_new_insulation_material, item: new Number(label: insulation_material_1_amount_of_new_insulation_material_lbl, min: 0, unit: "kg", order: ++order));
            igBuildingShell.Add(key: insulation_material_1_transport_to_building_by_truck, item: new Number(label: insulation_material_1_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: insulation_material_1_transport_to_building_by_train, item: new Number(label: insulation_material_1_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: insulation_material_1_transport_to_building_by_ferry, item: new Number(label: insulation_material_1_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Insulation material 2
            igBuildingShell.Add(key: change_insulation_material_2, item: new Checkbox(label: change_insulation_material_2_lbl, order: ++order));
            igBuildingShell.Add(key: insulation_material_2_life_of_product, item: new Number(label: insulation_material_2_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igBuildingShell.Add(key: insulation_material_2_type_of_insulation, item: new Select(label: insulation_material_2_type_of_insulation_lbl, options: type_of_insulation, order: ++order));
            igBuildingShell.Add(key: insulation_material_2_change_in_annual_heat_demand_due_to_insulation, item: new Number(label: insulation_material_2_change_in_annual_heat_demand_due_to_insulation_lbl, unit: "kWh/year", order: ++order));
            igBuildingShell.Add(key: insulation_material_2_amount_of_new_insulation_material, item: new Number(label: insulation_material_2_amount_of_new_insulation_material_lbl, min: 0, unit: "kg", order: ++order));
            igBuildingShell.Add(key: insulation_material_2_transport_to_building_by_truck, item: new Number(label: insulation_material_2_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: insulation_material_2_transport_to_building_by_train, item: new Number(label: insulation_material_2_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: insulation_material_2_transport_to_building_by_ferry, item: new Number(label: insulation_material_2_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Fascade System
            igBuildingShell.Add(key: change_fascade_system, item: new Checkbox(label: change_fascade_system_lbl, order: ++order));
            igBuildingShell.Add(key: fascade_system_life_of_product, item: new Number(label: fascade_system_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igBuildingShell.Add(key: fascade_system_type_fascade_system, item: new Select(label: fascade_system_type_of_fascade_system_lbl, options: type_of_fascade_system, order: ++order));
            igBuildingShell.Add(key: fascade_system_change_in_annual_heat_demand_due_to_fascade_system, item: new Number(label: fascade_system_change_in_annual_heat_demand_due_to_fascade_system_lbl, unit: "kWh/year", order: ++order));
            igBuildingShell.Add(key: fascade_system_area_of_new_fascade_system, item: new Number(label: fascade_system_area_of_new_fascade_system_lbl, min: 0, unit: "m\u00b2", order: ++order));
            igBuildingShell.Add(key: fascade_system_transport_to_building_by_truck, item: new Number(label: fascade_system_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: fascade_system_transport_to_building_by_train, item: new Number(label: fascade_system_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: fascade_system_transport_to_building_by_ferry, item: new Number(label: fascade_system_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Windows
            igBuildingShell.Add(key: change_windows, item: new Checkbox(label: change_windows_lbl, order: ++order));
            igBuildingShell.Add(key: windows_life_of_product, item: new Number(label: windows_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igBuildingShell.Add(key: windows_type_windows, item: new Select(label: windows_type_of_windows_lbl, options: type_of_windows, order: ++order));
            igBuildingShell.Add(key: windows_change_in_annual_heat_demand_due_to_windows, item: new Number(label: windows_change_in_annual_heat_demand_due_to_windows_lbl, unit: "kWh/year", order: ++order));
            igBuildingShell.Add(key: windows_area_of_new_windows, item: new Number(label: windows_area_of_new_windows_lbl, min: 0, unit: "m\u00b2", order: ++order));
            igBuildingShell.Add(key: windows_transport_to_building_by_truck, item: new Number(label: windows_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: windows_transport_to_building_by_train, item: new Number(label: windows_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: windows_transport_to_building_by_ferry, item: new Number(label: windows_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Doors
            igBuildingShell.Add(key: change_doors, item: new Checkbox(label: change_doors_lbl, order: ++order));
            igBuildingShell.Add(key: doors_life_of_product, item: new Number(label: doors_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igBuildingShell.Add(key: doors_type_doors, item: new Select(label: doors_type_of_doors_lbl, options: type_of_doors, order: ++order));
            igBuildingShell.Add(key: doors_change_in_annual_heat_demand_due_to_doors, item: new Number(label: doors_change_in_annual_heat_demand_due_to_doors_lbl, unit: "kWh/year", order: ++order));
            igBuildingShell.Add(key: doors_number_of_new_front_doors, item: new Number(label: doors_number_of_new_front_doors_lbl, min: 0, order: ++order));
            igBuildingShell.Add(key: doors_transport_to_building_by_truck, item: new Number(label: doors_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: doors_transport_to_building_by_train, item: new Number(label: doors_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igBuildingShell.Add(key: doors_transport_to_building_by_ferry, item: new Number(label: doors_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));



            return igBuildingShell;
        }

        InputGroup VentilationSystem(int ipgOrder = -1)
        {
            int order = 0;
            InputGroup igVentilationSystem = new InputGroup("Ventilation System", ++ipgOrder);

            // Ventilation ducts
            igVentilationSystem.Add(key: change_ventilation_ducts, item: new Checkbox(label: change_ventilation_ducts_lbl, order: ++order));
            igVentilationSystem.Add(key: ventilation_ducts_life_of_product, item: new Number(label: ventilation_ducts_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igVentilationSystem.Add(key: ventilation_ducts_type_of_material, item: new Select(label: ventilation_ducts_type_of_material_lbl, options: type_of_ventilation_ducts_material, order: ++order));
            igVentilationSystem.Add(key: ventilation_ducts_weight_of_ventilation_ducts, item: new Number(label: ventilation_ducts_weight_of_ventilation_ducts_lbl, unit: "kWh/year", order: ++order));
            igVentilationSystem.Add(key: ventilation_ducts_transport_to_building_by_truck, item: new Number(label: ventilation_ducts_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igVentilationSystem.Add(key: ventilation_ducts_transport_to_building_by_train, item: new Number(label: ventilation_ducts_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igVentilationSystem.Add(key: ventilation_ducts_transport_to_building_by_ferry, item: new Number(label: ventilation_ducts_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Airflow assembly
            igVentilationSystem.Add(key: change_airflow_assembly, item: new Checkbox(label: change_airflow_assembly_lbl, order: ++order));
            igVentilationSystem.Add(key: airflow_assembly_life_of_product, item: new Number(label: airflow_assembly_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igVentilationSystem.Add(key: airflow_assembly_type_of_airflow_assembly, item: new Select(label: airflow_assembly_type_of_airflow_assembly_lbl, options: type_of_airflow_assembly, order: ++order));
            igVentilationSystem.Add(key: airflow_assembly_design_airflow_exhaust_air, item: new Number(label: airflow_assembly_design_airflow_exhaust_air_lbl, unit: "kWh/year", order: ++order));
            igVentilationSystem.Add(key: airflow_assembly_transport_to_building_by_truck, item: new Number(label: airflow_assembly_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igVentilationSystem.Add(key: airflow_assembly_transport_to_building_by_train, item: new Number(label: airflow_assembly_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igVentilationSystem.Add(key: airflow_assembly_transport_to_building_by_ferry, item: new Number(label: airflow_assembly_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Air distribution housings and silencers
            igVentilationSystem.Add(key: change_air_distribution_housings_and_silencers, item: new Checkbox(label: change_air_distribution_housings_and_silencers_lbl, order: ++order));
            igVentilationSystem.Add(key: air_distribution_housings_and_silencers_life_of_product, item: new Number(label: air_distribution_housings_and_silencers_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igVentilationSystem.Add(key: air_distribution_housings_and_silencers_transport_to_building_by_truck, item: new Number(label: air_distribution_housings_and_silencers_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igVentilationSystem.Add(key: air_distribution_housings_and_silencers_transport_to_building_by_train, item: new Number(label: air_distribution_housings_and_silencers_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igVentilationSystem.Add(key: air_distribution_housings_and_silencers_transport_to_building_by_ferry, item: new Number(label: air_distribution_housings_and_silencers_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            //Ventilation renovation
            igVentilationSystem.Add(key: ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation, item: new Number(label: ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation_lbl, min: 0, unit: "kWh/year", order: ++order));
            igVentilationSystem.Add(key: ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation, item: new Number(label: ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation_lbl, min: 0, unit: "kWh/year", order: ++order));



            return igVentilationSystem;
        }

        InputGroup RadiatorsPipesElectricity(int ipgOrder = -1)
        {
            int order = 0;
            InputGroup igRadiatorsPipesElectricity = new InputGroup("Radiators, Pipes and Electricity", ++ipgOrder);

            // Radiators
            igRadiatorsPipesElectricity.Add(key: change_radiators, item: new Checkbox(label: change_radiators_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: radiators_life_of_product, item: new Number(label: radiators_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: radiators_type_of_radiators, item: new Select(label: radiators_type_of_radiators_lbl, options: type_of_radiators, order: ++order));
            igRadiatorsPipesElectricity.Add(key: radiators_weight_of_radiators, item: new Number(label: radiators_weight_of_radiators_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: radiators_transport_to_building_by_truck, item: new Number(label: radiators_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: radiators_transport_to_building_by_train, item: new Number(label: radiators_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: radiators_transport_to_building_by_ferry, item: new Number(label: radiators_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));
            
            // Piping System - Copper
            igRadiatorsPipesElectricity.Add(key: change_piping_copper, item: new Checkbox(label: change_piping_copper_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_copper_life_of_product, item: new Number(label: piping_copper_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_copper_weight_of_copper_pipes, item: new Number(label: piping_copper_weight_of_copper_pipes_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_copper_transport_to_building_by_truck, item: new Number(label: piping_copper_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_copper_transport_to_building_by_train, item: new Number(label: piping_copper_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_copper_transport_to_building_by_ferry, item: new Number(label: piping_copper_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));
            
            // Piping System - PEX
            igRadiatorsPipesElectricity.Add(key: change_piping_pex, item: new Checkbox(label: change_piping_pex_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pex_life_of_product, item: new Number(label: piping_pex_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pex_weight_of_pex_pipes, item: new Number(label: piping_pex_weight_of_pex_pipes_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pex_transport_to_building_by_truck, item: new Number(label: piping_pex_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pex_transport_to_building_by_train, item: new Number(label: piping_pex_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pex_transport_to_building_by_ferry, item: new Number(label: piping_pex_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - PP
            igRadiatorsPipesElectricity.Add(key: change_piping_pp, item: new Checkbox(label: change_piping_pp_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pp_life_of_product, item: new Number(label: piping_pp_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pp_weight_of_pp_pipes, item: new Number(label: piping_pp_weight_of_pp_pipes_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pp_transport_to_building_by_truck, item: new Number(label: piping_pp_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pp_transport_to_building_by_train, item: new Number(label: piping_pp_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_pp_transport_to_building_by_ferry, item: new Number(label: piping_pp_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - Cast Iron
            igRadiatorsPipesElectricity.Add(key: change_piping_cast_iron, item: new Checkbox(label: change_piping_cast_iron_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_cast_iron_life_of_product, item: new Number(label: piping_cast_iron_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_cast_iron_weight_of_cast_iron_pipes, item: new Number(label: piping_cast_iron_weight_of_cast_iron_pipes_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_cast_iron_transport_to_building_by_truck, item: new Number(label: piping_cast_iron_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_cast_iron_transport_to_building_by_train, item: new Number(label: piping_cast_iron_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_cast_iron_transport_to_building_by_ferry, item: new Number(label: piping_cast_iron_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - Galvanized Steel
            igRadiatorsPipesElectricity.Add(key: change_piping_galvanized_steel, item: new Checkbox(label: change_piping_galvanized_steel_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_galvanized_steel_life_of_product, item: new Number(label: piping_galvanized_steel_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_galvanized_steel_weight_of_galvanized_steel_pipes, item: new Number(label: piping_galvanized_steel_weight_of_galvanized_steel_pipes_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_galvanized_steel_transport_to_building_by_truck, item: new Number(label: piping_galvanized_steel_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_galvanized_steel_transport_to_building_by_train, item: new Number(label: piping_galvanized_steel_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_galvanized_steel_transport_to_building_by_ferry, item: new Number(label: piping_galvanized_steel_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - Relining
            igRadiatorsPipesElectricity.Add(key: change_piping_relining, item: new Checkbox(label: change_piping_relining_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_relining_life_of_product, item: new Number(label: piping_relining_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_relining_weight_of_relining_pipes, item: new Number(label: piping_relining_weight_of_relining_pipes_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_relining_transport_to_building_by_truck, item: new Number(label: piping_relining_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_relining_transport_to_building_by_train, item: new Number(label: piping_relining_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: piping_relining_transport_to_building_by_ferry, item: new Number(label: piping_relining_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Electrical wiring
            igRadiatorsPipesElectricity.Add(key: change_electrical_wiring, item: new Checkbox(label: change_electrical_wiring_lbl, order: ++order));
            igRadiatorsPipesElectricity.Add(key: electrical_wiring_life_of_product, item: new Number(label: electrical_wiring_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            igRadiatorsPipesElectricity.Add(key: electrical_wiring_weight_of_electrical_wiring, item: new Number(label: electrical_wiring_weight_of_electrical_wiring_lbl, unit: "kg", order: ++order));
            igRadiatorsPipesElectricity.Add(key: electrical_wiring_transport_to_building_by_truck, item: new Number(label: electrical_wiring_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: electrical_wiring_transport_to_building_by_train, item: new Number(label: electrical_wiring_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            igRadiatorsPipesElectricity.Add(key: electrical_wiring_transport_to_building_by_ferry, item: new Number(label: electrical_wiring_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            return igRadiatorsPipesElectricity;
        }


        // Simple version for one building

        InputSpecification GetInputSpecificationOneBuilding()
        {
            InputSpecification iSpec = new InputSpecification();
            
            // - ## Building Specific
            int order = 0;

            // Instructions
            string intstr = "";
            intstr = "Fill in the building specific data below. ";
            intstr += "Use the checkboxes to indicate what types of renovation procedures you want to perform for this alternative. ";
            intstr += "You need to fill in the  building properties as well as the parameters under checked checkboxes. ";
            intstr += "If this is the as-is step leave all checkboxes unchecked and fill in only the building properties. ";
            intstr += "If multiple buildings have common properties you may select all of them and assign them values simultaneously. ";
            InputGroup instructions = new InputGroup(label: intstr, order: ++order);
            iSpec.Add(key: "instructions", value: instructions);
            iSpec.Add(lca_calculation_period, new Number(label: "LCA calculation period", min: 1, unit: "years", order: ++order));

            // Building Common
            iSpec.Add(key: "properties", value: BuildingProperties(++order));

            // Heating System
            iSpec.Add(key: "heating_system", value: HeatingSystem(++order));

            // Building Shell
            iSpec.Add(key: "building_shell", value: BuildingShell(++order));

            // Ventilation System
            iSpec.Add(key: "ventilation_system", value: VentilationSystem(++order));

            // Radiators, pipes and electricity
            iSpec.Add(key: "radiators_pipes_and_electricity", value: RadiatorsPipesElectricity(++order));

            return iSpec;
        }

        void SetInputDataOneBuilding(Dictionary<string, Input> indata, ref CExcel exls)
        {
            // Single Building (simple)
            #region LCA Calculation Period
            String Key = lca_calculation_period;
            if (indata[Key] is Number)
            {
                var value = indata[Key] as Number;
                String cell = "C16";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    indata[Key].GetType(),
                    typeof(Number)));
            #endregion

            SetBuildingProperties(indata["properties"], ref exls);
            SetHeatingSystem(indata["heating_system"], ref exls);
        }
        //

        void SetBuildingProperties(Input input, ref CExcel exls)
        {
            if (!(input is InputGroup))
                throw new Exception("SetBuildingProperties: wrong input format!");

            InputGroup igBuildingCommon = input as InputGroup;

            Dictionary<string,Input> buildingCommonInputs = igBuildingCommon.GetInputs();
            String Key;

            // Inputs required in all cases
            #region Heated Area
            Key = heated_area;
            if (buildingCommonInputs[Key] is Number)
            {
                var value = buildingCommonInputs[Key] as Number;
                String cell = "C25";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}", 
                    buildingCommonInputs[Key].GetType(), 
                    typeof(Number)));
            #endregion

            #region Number of Apartments
            Key = nr_apartments;
            if (buildingCommonInputs[Key] is Number)
            {
                var value = buildingCommonInputs[Key] as Number;
                String cell = "C26";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    buildingCommonInputs[Key].GetType(),
                    typeof(Number)));
            #endregion
            
            #region Heat Source Before
            Key = heat_source_before;
            if (buildingCommonInputs[Key] is Select)
            {
                var value = buildingCommonInputs[Key] as Select;
                String cell = "C93";
                if (value.SelectedIndex() >= 0)
                {
                    if (!exls.SetCellValue("Indata", cell, value.SelectedIndex()))
                        throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value.SelectedIndex()));
                }
                else
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value.SelectedIndex()));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    buildingCommonInputs[Key].GetType(),
                    typeof(Select)));
            #endregion

            #region Heat Source After
            Key = heat_source_after;
            if (buildingCommonInputs[Key] is Select)
            {
                var value = buildingCommonInputs[Key] as Select;
                String cell = "C93";
                if (value.SelectedIndex() >= 0)
                {
                    if (!exls.SetCellValue("Indata", cell, value.SelectedIndex()))
                        throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value.SelectedIndex()));
                }
                else
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value.SelectedIndex()));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    buildingCommonInputs[Key].GetType(),
                    typeof(Select)));
            #endregion

            // If district heating is used (before/after renovation)
            #region Global warming potential of district heating
            Key = gwp_district;
            if (buildingCommonInputs[Key] is Number)
            {
                var value = buildingCommonInputs[Key] as Number;
                String cell = "C26";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    buildingCommonInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region "Primary energy use of district heating
            Key = peu_district;
            if (buildingCommonInputs[Key] is Number)
            {
                var value = buildingCommonInputs[Key] as Number;
                String cell = "C26";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    buildingCommonInputs[Key].GetType(),
                    typeof(Number)));
            #endregion
        }

        void SetHeatingSystem(Input input, ref CExcel exls)
        {
            if (!(input is InputGroup))
                throw new Exception("SetHeatingSystem: wrong input format!");

            InputGroup igHeatingSystem = input as InputGroup;

            Dictionary<string, Input> heatingSystemInputs = igHeatingSystem.GetInputs();
            String Key;

            // Change Heating System
            #region Change Heating System
            Key = change_heating_system;
            if (heatingSystemInputs[Key] is Checkbox)
            {
                var value = heatingSystemInputs[Key] as Checkbox;
                String cell = "C99";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Checkbox)));
            #endregion

            #region Annual Heat Demand After Renovation
            Key = ahd_after_renovation;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Heating System: Life of Product
            Key = heating_system_life_of_product;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C100";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Design Capacity
            Key = design_capacity;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C103";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Weight of boiler/heat pump/district heating substation
            Key = weight_of_bhd;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C103";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region For geothermal heat pump: Depth of borehole
            Key = depth_of_borehole;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C103";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Transport to building by truck
            Key = heating_system_transport_to_building_truck;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C103";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Transport to building by train
            Key = heating_system_transport_to_building_train;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C103";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Transport to building by ferry
            Key = heating_system_transport_to_building_ferry;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C103";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            // Change Circulation Pump
            #region Change Circulation Pump
            Key = change_circulationpump_in_heating_system;
            if (heatingSystemInputs[Key] is Checkbox)
            {
                var value = heatingSystemInputs[Key] as Checkbox;
                String cell = "C99";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Checkbox)));
            #endregion

            #region Circulation Pump: Life of Product
            Key = circulationpump_life_of_product;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Design pressure head
            Key = design_pressure_head;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Design flow rate
            Key = design_flow_rate;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Type of flow control in heating system
            Key = type_of_control_in_heating_system;
            if (heatingSystemInputs[Key] is Select)
            {
                var value = heatingSystemInputs[Key] as Select;
                String cell = "C93";
                if (value.SelectedIndex() >= 0)
                {
                    if (!exls.SetCellValue("Indata", cell, value.SelectedIndex()))
                        throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value.SelectedIndex()));
                }
                else
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value.SelectedIndex()));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Select)));
            #endregion

            #region Weight
            Key = weight;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Transport to building by truck
            Key = circulationpump_transport_to_building_truck;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Transport to building by train
            Key = circulationpump_transport_to_building_train;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

            #region Transport to building by ferry
            Key = circulationpump_transport_to_building_ferry;
            if (heatingSystemInputs[Key] is Number)
            {
                var value = heatingSystemInputs[Key] as Number;
                String cell = "C95";
                if (!exls.SetCellValue("Indata", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            }
            else
                throw new Exception(String.Format("Could not set cell, data in the wrong format. {0} instead of {1}",
                    heatingSystemInputs[Key].GetType(),
                    typeof(Number)));
            #endregion

        }

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            if(!inputSpecifications.ContainsKey(kpiId))
                throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.", kpiId));

            return inputSpecifications[kpiId];
        }

        protected override Outputs CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls)
        {
            Outputs outputs = new Outputs();

            SetInputDataOneBuilding(indata, ref exls);


            switch (kpiId)
            {
                case kpi_gwp:
                    var cgwp = exls.GetCellValue("Indata", "C31"); //Change of global warming potential
                    outputs.Add(new Kpi(cgwp, "Change of global warming potential", "tonnes CO2 eq"));
                    break;
                case kpi_peu:
                    var cpeu = exls.GetCellValue("Indata", "C32"); //Change of primary energy use                    
                    outputs.Add(new Kpi(cpeu, "Change of primary energy use", "MWh"));                    
                    break;
                default:
                    throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
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

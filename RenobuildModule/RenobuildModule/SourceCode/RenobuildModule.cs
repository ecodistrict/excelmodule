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

        Options electricity_mix_opts;
        Options heat_sources;
        Options type_of_flow_control_in_heating_system_opts;

        Options type_of_insulation;
        Options type_of_facade_system;
        Options type_of_windows;
        Options type_of_doors;

        Options type_of_ventilation_ducts_material;
        Options type_of_airflow_assembly;

        Options type_of_radiators;

        void DefineElectricityMix()
        {
            try
            {
                electricity_mix_opts = new Options();
                electricity_mix_opts.Add(new Option(value: "em_sweden", label: "Sweden"));
                electricity_mix_opts.Add(new Option(value: "em_netherlands", label: "Netherlands"));
                electricity_mix_opts.Add(new Option(value: "em_spain", label: "Spain"));
                electricity_mix_opts.Add(new Option(value: "em_poland", label: "Poland"));
                electricity_mix_opts.Add(new Option(value: "em_belgium", label: "Belgium"));
            }
            catch (Exception ex)
            {
                RenobuildModule_ErrorRaised(this, ex);
            }
        }
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
            catch (Exception ex)
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
                type_of_facade_system = new Options();
                type_of_facade_system.Add(new Option(value: @"A\8-15mm\Non ventilated\EPS\200mm", label: @"A\8-15mm\Non ventilated\EPS\200mm"));
                type_of_facade_system.Add(new Option(value: @"B\4-8mm\Ventilated\Rock wool\50mm", label: @"B\4-8mm\Ventilated\Rock wool\50mm"));
                type_of_facade_system.Add(new Option(value: @"B\4-8mm\Ventilated\Rock wool\80mm", label: @"B\4-8mm\Ventilated\Rock wool\80mm"));
                type_of_facade_system.Add(new Option(value: @"B\4-8mm\Ventilated\Rock wool\100mm", label: @"B\4-8mm\Ventilated\Rock wool\100mm"));
                type_of_facade_system.Add(new Option(value: @"C\8-12mm\Non ventilated\EPS\50mm", label: @"C\8-12mm\Non ventilated\EPS\50mm"));
                type_of_facade_system.Add(new Option(value: @"C\8-12mm\Non ventilated\EPS\80mm", label: @"C\8-12mm\Non ventilated\EPS\80mm"));
                type_of_facade_system.Add(new Option(value: @"C\8-12mm\Non ventilated\EPS\100mm", label: @"C\8-12mm\Non ventilated\EPS\100mm"));
                type_of_facade_system.Add(new Option(value: @"D\20-mm\Non ventilated\Rock wool\50mm", label: @"D\20-mm\Non ventilated\Rock wool\50mm"));
                type_of_facade_system.Add(new Option(value: @"D\20-mm\Non ventilated\Rock wool\80mm", label: @"D\20-mm\Non ventilated\Rock wool\80mm"));
                type_of_facade_system.Add(new Option(value: @"D\20-mm\Non ventilated\Rock wool\100mm", label: @"D\20-mm\Non ventilated\Rock wool\100mm"));
                type_of_facade_system.Add(new Option(value: @"E\10-15mm\Non ventilated\Rock wool\50mm", label: @"E\10-15mm\Non ventilated\Rock wool\50mm"));
                type_of_facade_system.Add(new Option(value: @"E\10-15mmNon ventilated\Rock wool\80mm", label: @"E\10-15mmNon ventilated\Rock wool\80mm"));
                type_of_facade_system.Add(new Option(value: @"E\10-15mm\Non ventilated\Rock wool\100mm", label: @"E\10-15mm\Non ventilated\Rock wool\100mm"));
                type_of_facade_system.Add(new Option(value: @"E\10-15mm\Non ventilated\Rock wool, PIR\50+150mm", label: @"E\10-15mm\Non ventilated\Rock wool, PIR\50+150mm"));
                type_of_facade_system.Add(new Option(value: @"F\4-8mm\Ventilated\Rock wool\80mm", label: @"F\4-8mm\Ventilated\Rock wool\80mm"));
                type_of_facade_system.Add(new Option(value: @"F\4-8mm\Ventilated\Rock wool\100mm", label: @"F\4-8mm\Ventilated\Rock wool\100mm"));
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
                inputSpecifications.Add(kpi_gwp, GetInputSpecificationGeoJson());
                inputSpecifications.Add(kpi_peu, GetInputSpecificationGeoJson());
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
        string electricity_mix = "electricity_mix";

        #endregion

        #region Building specific Properties
        // Building specific Properties
        string buildings = "buildings";
        #region Building Common
        //Building Common
        // Inputs required in all cases
        string heated_area = "HEATED_AREA";
        string heated_area_lbl = "Heated area";
        string nr_apartments = "NUMBER_OF_APARTMENTS";
        string nr_apartments_lbl = "Number of apartments";
        string heat_source_before = "HEAT_SOURCE_BEFORE";
        string heat_source_before_lbl = "Heat source before renovation";
        string heat_source_after = "HEAT_SOURCE_AFTER";
        string heat_source_after_lbl = "Heat source after renovation";
        // If district heating is used (before/after renovation)
        string gwp_district = "GWP_OF_DISTRICT_HEATING";
        string gwp_district_lbl = "Global warming potential of district heating. Required if any building uses district heating before or after renovation. Impact per unit energy delivered to building, i.e. including distribution losses.";
        string peu_district = "PRIMARY_ENERGY_USE_OF_DISTRICT_HEATING";
        string peu_district_lbl = "Primary energy factor of district heating. Required if any building uses district heating before or after renovation. Impact per unit energy delivered to building, i.e. including distribution losses.";
        #endregion

        #region Heating System
        //Heating System
        // Change Heating System
        string change_heating_system = "HEATING_SYSTEM__CHANGE_HEATING_SYSTEM";
        string change_heating_system_lbl = "Replace building heating system";
        string ahd_after_renovation = "HEATING_SYSTEM__AHD_AFTER_RENOVATION";
        string ahd_after_renovation_lbl = "Annual heat demand after renovation";
        string heating_system_life_of_product = "HEATING_SYSTEM__LIFE_OF_PRODUCT";
        string heating_system_life_of_product_lbl = "Life of product (Practical time of life of the products and materials used)";
        string design_capacity = "HEATING_SYSTEM__DESIGN_CAPACITY";
        string design_capacity_lbl = "Design capacity (Required for pellets boiler and oil boiler)";
        string weight_of_bhd = "HEATING_SYSTEM__WEIGHT_OF_BHD";
        string weight_of_bhd_lbl = "Weight of boiler/heat pump/district heating substation (Required except for direct electricity heating)";
        string depth_of_borehole = "HEATING_SYSTEM__DEPTH_OF_BORE_HOLE";
        string depth_of_borehole_lbl = "Depth of bore hole (For geothermal heat pump)";
        string heating_system_transport_to_building_truck = "HEATING_SYSTEM__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string heating_system_transport_to_building_truck_lbl = "Transport to building by truck (Distance from production site to building)";
        string heating_system_transport_to_building_train = "HEATING_SYSTEM__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string heating_system_transport_to_building_train_lbl = "Transport to building by train (Distance from production site to building)";
        string heating_system_transport_to_building_ferry = "HEATING_SYSTEM__TRANSPORT_TO_BUILDING_BY_FERRY";
        string heating_system_transport_to_building_ferry_lbl = "Transport to building by ferry (Distance from production site to building)";

        // Change Circulation Pump
        string change_circulationpump_in_heating_system = "PUMP__CHANGE_PUMP";
        string change_circulationpump_in_heating_system_lbl = "Replace circulation pump in building heating system";
        string circulationpump_life_of_product = "PUMP__LIFE_OF_PRODUCT";
        string circulationpump_life_of_product_lbl = "Practical time of life of the products and materials used";
        string design_pressure_head = "PUMP__DESIGN_PRESSURE_HEAD";
        string design_pressure_head_lbl = "Design pressure head";
        string design_flow_rate = "PUMP__DESIGN_FLOW_RATE";
        string design_flow_rate_lbl = "Design flow rate";
        string type_of_control_in_heating_system = "PUMP__TYPE_OF_FLOW_CONTROL";
        string type_of_control_in_heating_system_lbl = "Type of flow control in heating system";
        string weight = "PUMP__WEIGHT";
        string weight_lbl = "Weight of new pump";
        string circulationpump_transport_to_building_truck = "PUMP__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string circulationpump_transport_to_building_truck_lbl = "Transport to building by truck (Distance from production site to building)";
        string circulationpump_transport_to_building_train = "PUMP__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string circulationpump_transport_to_building_train_lbl = "Transport to building by train (Distance from production site to building)";
        string circulationpump_transport_to_building_ferry = "PUMP__TRANSPORT_TO_BUILDING_BY_FERRY";
        string circulationpump_transport_to_building_ferry_lbl = "Transport to building by ferry (Distance from production site to building)";
        #endregion
        
        #region Building Shell
        //Building Shell
        // Insulation material 1
        string change_insulation_material_1 = "INSULATION_MATERIAL_ONE__CHANGE";
        string change_insulation_material_1_lbl = "Use insulation material 1";
        string insulation_material_1_life_of_product = "INSULATION_MATERIAL_ONE__LIFE_OF_PRODUCT";
        string insulation_material_1_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string insulation_material_1_type_of_insulation = "INSULATION_MATERIAL_ONE__TYPE_OF_INSULATION";
        string insulation_material_1_type_of_insulation_lbl = "Type of insulation";
        string insulation_material_1_change_in_annual_heat_demand_due_to_insulation = "INSULATION_MATERIAL_ONE__CHANGE_IN_AHD_DUE_TO_INSULATION";
        string insulation_material_1_change_in_annual_heat_demand_due_to_insulation_lbl = "Change in annual heat demand due to insulation (an energy saving is given as a negative value)";
        string insulation_material_1_amount_of_new_insulation_material = "INSULATION_MATERIAL_ONE__AMOUNT_OF_NEW_INSULATION_MATERIAL";
        string insulation_material_1_amount_of_new_insulation_material_lbl = "Amount of new insulation material (required if renovation includes new insulation material)";
        string insulation_material_1_transport_to_building_by_truck = "INSULATION_MATERIAL_ONE__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string insulation_material_1_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string insulation_material_1_transport_to_building_by_train = "INSULATION_MATERIAL_ONE__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string insulation_material_1_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string insulation_material_1_transport_to_building_by_ferry = "INSULATION_MATERIAL_ONE__TRANSPORT_TO_BUILDING_BY_FERRY";
        string insulation_material_1_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Insulation material 2
        string change_insulation_material_2 = "INSULATION_MATERIAL_TWO__CHANGE";
        string change_insulation_material_2_lbl = "Use insulation material 2";
        string insulation_material_2_life_of_product = "INSULATION_MATERIAL_TWO__LIFE_OF_PRODUCT";
        string insulation_material_2_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string insulation_material_2_type_of_insulation = "INSULATION_MATERIAL_TWO__TYPE_OF_INSULATION";
        string insulation_material_2_type_of_insulation_lbl = "Type of insulation";
        string insulation_material_2_change_in_annual_heat_demand_due_to_insulation = "INSULATION_MATERIAL_TWO__CHANGE_IN_AHD_DUE_TO_INSULATION";
        string insulation_material_2_change_in_annual_heat_demand_due_to_insulation_lbl = "Change in annual heat demand due to insulation (an energy saving is given as a negative value)";
        string insulation_material_2_amount_of_new_insulation_material = "INSULATION_MATERIAL_TWO__AMOUNT_OF_NEW_INSULATION_MATERIAL";
        string insulation_material_2_amount_of_new_insulation_material_lbl = "Amount of new insulation material (required if renovation includes new insulation material)";
        string insulation_material_2_transport_to_building_by_truck = "INSULATION_MATERIAL_TWO__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string insulation_material_2_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string insulation_material_2_transport_to_building_by_train = "INSULATION_MATERIAL_TWO__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string insulation_material_2_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string insulation_material_2_transport_to_building_by_ferry = "INSULATION_MATERIAL_TWO__TRANSPORT_TO_BUILDING_BY_FERRY";
        string insulation_material_2_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // facade system
        string change_facade_system = "FACADE__CHANGE";
        string change_facade_system_lbl = "Change facade";
        string facade_system_life_of_product = "FACADE__LIFE_OF_PRODUCT";
        string facade_system_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string facade_system_type_facade_system = "FACADE__TYPE_OF_FACADE_SYSTEM";
        string facade_system_type_of_facade_system_lbl = "Type of facade system";
        string facade_system_change_in_annual_heat_demand_due_to_facade_system = "FACADE__CHANGE_IN_AHD_DUE_TO_FACADE_SYSTEM";
        string facade_system_change_in_annual_heat_demand_due_to_facade_system_lbl = "Change in annual heat demand due to facade system (an energy saving is given as a negative value)";
        string facade_system_area_of_new_facade_system = "FACADE__AREA_OF_NEW_FACADE_SYSTEM";
        string facade_system_area_of_new_facade_system_lbl = "Area of new facade system (required if renovation includes new facade system)";
        string facade_system_transport_to_building_by_truck = "FACADE__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string facade_system_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string facade_system_transport_to_building_by_train = "FACADE__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string facade_system_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string facade_system_transport_to_building_by_ferry = "FACADE__TRANSPORT_TO_BUILDING_BY_FERRY";
        string facade_system_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";


        // Windows
        string change_windows = "WINDOWS__CHANGE";
        string change_windows_lbl = "Change windows";
        string windows_life_of_product = "WINDOWS__LIFE_OF_PRODUCT";
        string windows_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string windows_type_windows = "WINDOWS__TYPE_OF_WINDOWS";
        string windows_type_of_windows_lbl = "Material in frame";
        string windows_change_in_annual_heat_demand_due_to_windows = "WINDOWS__CHANGE_IN_AHD_DUE_TO_WINDOWS";
        string windows_change_in_annual_heat_demand_due_to_windows_lbl = "Change in annual heat demand due to windows (an energy saving is given as a negative value)";
        string windows_area_of_new_windows = "WINDOWS__AREA_OF_NEW_WINDOWS";
        string windows_area_of_new_windows_lbl = "Area of windows (required if renovation includes new windows)";
        string windows_transport_to_building_by_truck = "WINDOWS__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string windows_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string windows_transport_to_building_by_train = "WINDOWS__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string windows_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string windows_transport_to_building_by_ferry = "WINDOWS__TRANSPORT_TO_BUILDING_BY_FERRY";
        string windows_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Doors
        string change_doors = "DOORS__CHANGE";
        string change_doors_lbl = "Change doors";
        string doors_life_of_product = "DOORS__LIFE_OF_PRODUCT";
        string doors_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string doors_type_doors = "DOORS__TYPE_OF_DOORS";
        string doors_type_of_doors_lbl = "Type of doors";
        string doors_change_in_annual_heat_demand_due_to_doors = "DOORS__CHANGE_IN_AHD_DUE_TO_DOORS";
        string doors_change_in_annual_heat_demand_due_to_doors_lbl = "Change in annual heat demand due to doors (an energy saving is given as a negative value)";
        string doors_number_of_new_front_doors = "DOORS__NUMBER_OF_NEW_FRONT_DOORS";
        string doors_number_of_new_front_doors_lbl = "Number of new front doors (required if renovation includes new doors)";
        string doors_transport_to_building_by_truck = "DOORS__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string doors_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string doors_transport_to_building_by_train = "DOORS__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string doors_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string doors_transport_to_building_by_ferry = "DOORS__TRANSPORT_TO_BUILDING_BY_FERRY";
        string doors_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";
        #endregion

        #region Ventilation
        // Ventilation
        // Ventilation ducts
        string change_ventilation_ducts = "VENTILATION_DUCTS__CHANGE";
        string change_ventilation_ducts_lbl = "Change ventilation ducts";
        string ventilation_ducts_life_of_product = "VENTILATION_DUCTS__LIFE_OF_PRODUCT";
        string ventilation_ducts_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string ventilation_ducts_type_of_material = "VENTILATION_DUCTS__MATERIAL_OF_VENTILATION_DUCTS";
        string ventilation_ducts_type_of_material_lbl = "Material in ventilation ducts";
        string ventilation_ducts_weight_of_ventilation_ducts = "VENTILATION_DUCTS__WEIGHT_OF_VENTILATION_DUCTS";
        string ventilation_ducts_weight_of_ventilation_ducts_lbl = "Weight of ventilation ducts (Required if renovation includes new ventilation ducts)";
        string ventilation_ducts_transport_to_building_by_truck = "VENTILATION_DUCTS__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string ventilation_ducts_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string ventilation_ducts_transport_to_building_by_train = "VENTILATION_DUCTS__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string ventilation_ducts_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string ventilation_ducts_transport_to_building_by_ferry = "VENTILATION_DUCTS__TRANSPORT_TO_BUILDING_BY_FERRY";
        string ventilation_ducts_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Airflow assembly
        string change_airflow_assembly = "AIR_FLOW_ASSEMBLY__CHANGE";
        string change_airflow_assembly_lbl = "Change airflow assembly";
        string airflow_assembly_life_of_product = "AIR_FLOW_ASSEMBLY__LIFE_OF_PRODUCT";
        string airflow_assembly_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string airflow_assembly_type_of_airflow_assembly = "AIR_FLOW_ASSEMBLY__TYPE_OF_AIR_FLOW_ASSEMBLY";
        string airflow_assembly_type_of_airflow_assembly_lbl = "Type of airflow assembly";
        string airflow_assembly_design_airflow_exhaust_air = "AIR_FLOW_ASSEMBLY__DESIGN_AIR_FLOW";
        string airflow_assembly_design_airflow_exhaust_air_lbl = "Design airflow (exhaust air)";
        string airflow_assembly_transport_to_building_by_truck = "AIR_FLOW_ASSEMBLY__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string airflow_assembly_transport_to_building_by_truck_lbl = "Transport of airflow assembly to building by truck (distance from production site to building)";
        string airflow_assembly_transport_to_building_by_train = "AIR_FLOW_ASSEMBLY__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string airflow_assembly_transport_to_building_by_train_lbl = "Transport of airflow assembly to building by train (distance from production site to building)";
        string airflow_assembly_transport_to_building_by_ferry = "AIR_FLOW_ASSEMBLY__TRANSPORT_TO_BUILDING_BY_FERRY";
        string airflow_assembly_transport_to_building_by_ferry_lbl = "Transport of airflow assembly to building by ferry (distance from production site to building)";

        // Air distribution housings and silencer
        string change_air_distribution_housings_and_silencers = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__CHANGE";
        string change_air_distribution_housings_and_silencers_lbl = "Change air distribution housings and silencers";
        string air_distribution_housings_and_silencers_number_of_distribution_housings = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__NUMBER_OF_NEW_HOUSINGS";
        string air_distribution_housings_and_silencers_number_of_distribution_housings_lbl = "Number of air distribution housings";
        string air_distribution_housings_and_silencers_life_of_product = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__LIFE_OF_PRODUCT";
        string air_distribution_housings_and_silencers_life_of_product_lbl = "Life of air distribution housings and silencers (practical time of life of the products and materials used)";
        string air_distribution_housings_and_silencers_transport_to_building_by_truck = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string air_distribution_housings_and_silencers_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string air_distribution_housings_and_silencers_transport_to_building_by_train = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string air_distribution_housings_and_silencers_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string air_distribution_housings_and_silencers_transport_to_building_by_ferry = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__TRANSPORT_TO_BUILDING_BY_FERRY";
        string air_distribution_housings_and_silencers_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        //Ventilation renovation
        string ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation = "VENTILATION_SYSTEM__CHANGE_IN_AHD_DUE_TO_VENTILATION_SYSTEM_RENOVATION";
        string ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation_lbl = "Change in annual heat demand due ventilation systems renovation (an energy saving is given as a negative value)";
        string ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation = "VENTILATION_SYSTEM__CHANGE_IN_AED_DUE_TO_VENTILATION_SYSTEM_RENOVATION";
        string ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation_lbl = "Change in annual electricity demand due ventilation systems renovation (an energy saving is given as a negative value)";

        #endregion
        
        string change_in_ahd_due_to_renovations_of_bshell_ventilation_pump = "CHANGE_IN_AHD_DUE_TO_RENOVATIONS";
        string change_in_ahd_due_to_renovations_of_bshell_ventilation_pump_lbl = "Change in annual heat demand";
        string change_in_aed_due_to_renovations_of_bshell_ventilation_pump = "CHANGE_IN_AED_DUE_TO_RENOVATIONS";
        string change_in_aed_due_to_renovations_of_bshell_ventilation_pump_lbl = "Change in annual energy demand";

        #region Radiators, pipes and electricity
        // Radiators, pipes and electricity
        // Radiators
        string change_radiators = "RADIATORS__CHANGE";
        string change_radiators_lbl = "Change radiators";
        string radiators_life_of_product = "RADIATORS__LIFE_OF_PRODUCT";
        string radiators_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string radiators_type_of_radiators = "RADIATORS__TYPE_OF_RADIATORS";
        string radiators_type_of_radiators_lbl = "Type of radiators";
        string radiators_weight_of_radiators = "RADIATORS__WEIGHT_OF_NEW_RADIATORS";
        string radiators_weight_of_radiators_lbl = "Weight of new radiators";
        string radiators_transport_to_building_by_truck = "RADIATORS__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string radiators_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string radiators_transport_to_building_by_train = "RADIATORS__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string radiators_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string radiators_transport_to_building_by_ferry = "RADIATORS__TRANSPORT_TO_BUILDING_BY_FERRY";
        string radiators_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Copper
        string change_piping_copper = "PIPING_SYSTEM_COPPER__CHANGE";
        string change_piping_copper_lbl = "Change copper pipes";
        string piping_copper_life_of_product = "PIPING_SYSTEM_COPPER__LIFE_OF_PRODUCT";
        string piping_copper_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_copper_weight_of_copper_pipes = "PIPING_SYSTEM_COPPER__WEIGHT_OF_NEW_PIPES";
        string piping_copper_weight_of_copper_pipes_lbl = "Weight of new pipes";
        string piping_copper_transport_to_building_by_truck = "PIPING_SYSTEM_COPPER__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string piping_copper_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_copper_transport_to_building_by_train = "PIPING_SYSTEM_COPPER__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string piping_copper_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_copper_transport_to_building_by_ferry = "PIPING_SYSTEM_COPPER__TRANSPORT_TO_BUILDING_BY_FERRY";
        string piping_copper_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - PEX
        string change_piping_pex = "PIPING_SYSTEM_PEX__CHANGE";
        string change_piping_pex_lbl = "Change PEX pipes";
        string piping_pex_life_of_product = "PIPING_SYSTEM_PEX__LIFE_OF_PRODUCT";
        string piping_pex_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_pex_weight_of_pex_pipes = "PIPING_SYSTEM_PEX__WEIGHT_OF_NEW_PIPES";
        string piping_pex_weight_of_pex_pipes_lbl = "Weight of new pipes";
        string piping_pex_transport_to_building_by_truck = "PIPING_SYSTEM_PEX__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string piping_pex_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_pex_transport_to_building_by_train = "PIPING_SYSTEM_PEX__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string piping_pex_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_pex_transport_to_building_by_ferry = "PIPING_SYSTEM_PEX__TRANSPORT_TO_BUILDING_BY_FERRY";
        string piping_pex_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - PP
        string change_piping_pp = "PIPING_SYSTEM_PP__CHANGE";
        string change_piping_pp_lbl = "Change PP pipes";
        string piping_pp_life_of_product = "PIPING_SYSTEM_PP__LIFE_OF_PRODUCT";
        string piping_pp_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_pp_weight_of_pp_pipes = "PIPING_SYSTEM_PP__WEIGHT_OF_NEW_PIPES";
        string piping_pp_weight_of_pp_pipes_lbl = "Weight of new pipes";
        string piping_pp_transport_to_building_by_truck = "PIPING_SYSTEM_PP__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string piping_pp_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_pp_transport_to_building_by_train = "PIPING_SYSTEM_PP__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string piping_pp_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_pp_transport_to_building_by_ferry = "PIPING_SYSTEM_PP__TRANSPORT_TO_BUILDING_BY_FERRY";
        string piping_pp_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Cast Iron
        string change_piping_cast_iron = "PIPING_SYSTEM_CAST_IRON__CHANGE";
        string change_piping_cast_iron_lbl = "Change cast iron pipes";
        string piping_cast_iron_life_of_product = "PIPING_SYSTEM_CAST_IRON__LIFE_OF_PRODUCT";
        string piping_cast_iron_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_cast_iron_weight_of_cast_iron_pipes = "PIPING_SYSTEM_CAST_IRON__WEIGHT_OF_NEW_PIPES";
        string piping_cast_iron_weight_of_cast_iron_pipes_lbl = "Weight of new pipes";
        string piping_cast_iron_transport_to_building_by_truck = "PIPING_SYSTEM_CAST_IRON__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string piping_cast_iron_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_cast_iron_transport_to_building_by_train = "PIPING_SYSTEM_CAST_IRON__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string piping_cast_iron_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_cast_iron_transport_to_building_by_ferry = "PIPING_SYSTEM_CAST_IRON__TRANSPORT_TO_BUILDING_BY_FERRY";
        string piping_cast_iron_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Galvanized Steel
        string change_piping_galvanized_steel = "PIPING_SYSTEM_GALVANISED_STEEL__CHANGE";
        string change_piping_galvanized_steel_lbl = "Change galvanized steel pipes";
        string piping_galvanized_steel_life_of_product = "PIPING_SYSTEM_GALVANISED_STEEL__LIFE_OF_PRODUCT";
        string piping_galvanized_steel_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_galvanized_steel_weight_of_galvanized_steel_pipes = "PIPING_SYSTEM_GALVANISED_STEEL__WEIGHT_OF_NEW_PIPES";
        string piping_galvanized_steel_weight_of_galvanized_steel_pipes_lbl = "Weight of new pipes";
        string piping_galvanized_steel_transport_to_building_by_truck = "PIPING_SYSTEM_GALVANISED_STEEL__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string piping_galvanized_steel_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_galvanized_steel_transport_to_building_by_train = "PIPING_SYSTEM_GALVANISED_STEEL__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string piping_galvanized_steel_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_galvanized_steel_transport_to_building_by_ferry = "PIPING_SYSTEM_GALVANISED_STEEL__TRANSPORT_TO_BUILDING_BY_FERRY";
        string piping_galvanized_steel_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Piping System - Relining
        string change_piping_relining = "PIPING_SYSTEM_RELINING__CHANGE";
        string change_piping_relining_lbl = "Relining of pipes";
        string piping_relining_life_of_product = "PIPING_SYSTEM_RELINING__LIFE_OF_PRODUCT";
        string piping_relining_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string piping_relining_weight_of_relining_pipes = "PIPING_SYSTEM_RELINING__WEIGHT_OF_NEW_PIPES";
        string piping_relining_weight_of_relining_pipes_lbl = "Weight of new pipes";
        string piping_relining_transport_to_building_by_truck = "PIPING_SYSTEM_RELINING__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string piping_relining_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string piping_relining_transport_to_building_by_train = "PIPING_SYSTEM_RELINING__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string piping_relining_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string piping_relining_transport_to_building_by_ferry = "PIPING_SYSTEM_RELINING__TRANSPORT_TO_BUILDING_BY_FERRY";
        string piping_relining_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        // Electrical wiring
        string change_electrical_wiring = "ELECTRICAL_WIRING__CHANGE";
        string change_electrical_wiring_lbl = "Replace electrical wiring";
        string electrical_wiring_life_of_product = "ELECTRICAL_WIRING__LIFE_OF_PRODUCT";
        string electrical_wiring_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        string electrical_wiring_weight_of_electrical_wiring = "ELECTRICAL_WIRING__WEIGHT_OF_NEW_WIRES";
        string electrical_wiring_weight_of_electrical_wiring_lbl = "Weight of new wires";
        string electrical_wiring_transport_to_building_by_truck = "ELECTRICAL_WIRING__TRANSPORT_TO_BUILDING_BY_TRUCK";
        string electrical_wiring_transport_to_building_by_truck_lbl = "Transport to building by truck (distance from production site to building)";
        string electrical_wiring_transport_to_building_by_train = "ELECTRICAL_WIRING__TRANSPORT_TO_BUILDING_BY_TRAIN";
        string electrical_wiring_transport_to_building_by_train_lbl = "Transport to building by train (distance from production site to building)";
        string electrical_wiring_transport_to_building_by_ferry = "ELECTRICAL_WIRING__TRANSPORT_TO_BUILDING_BY_FERRY";
        string electrical_wiring_transport_to_building_by_ferry_lbl = "Transport to building by ferry (distance from production site to building)";

        #endregion

        #endregion

        #endregion

        #endregion

        public RenobuildModule()
        {
            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "Renobuild";

            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_gwp, kpi_peu };

            //Error handler
            this.ErrorRaised += RenobuildModule_ErrorRaised;

            //Notification
            this.StatusMessage += RenobuildModule_StatusMessage;

            //Define parameter options
            DefineElectricityMix();
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
            string description = "Building specific properties (Use the geojson-upload functionality below the map in order change renovation options for your buildings. You can select one or more buildings at the time by clicking on them, when you are finished with the selected building(s) press OK for the input sheet and continue selecting other buildings. When you have supplied all data scroll all the way down and press OK.)";

            iSpec.Add("buildingProperties", new InputGroup(label: description, order: 2));
            iSpec.Add(buildings, BuildingSpecificSpecGeoJson());

            return iSpec;
        }

        InputGroup CommonSpec()
        {
            int order = 0;

            // - ## Common Properties
            InputGroup commonProp = new InputGroup(label: "Common properties", order: 1);
            commonProp.Add(lca_calculation_period, new Number(label: "LCA calculation period", min: 1, unit: "years", order: ++order));
            commonProp.Add(electricity_mix, new Select(label: "Electricity mix", options: electricity_mix_opts, order: ++order));
            // If district heating is used (before/after renovation)
            commonProp.Add(key: gwp_district, item: new Number(label: gwp_district_lbl, min: 0, unit: "g CO2 eq/kWh", order: ++order));
            commonProp.Add(key: peu_district, item: new Number(label: peu_district_lbl, min: 0, unit: "kWh/kWh", order: ++order));

            return commonProp;
        }

        GeoJson BuildingSpecificSpecGeoJson()
        {
            // - ## Building Specific
            GeoJson buildning_specific_data = new GeoJson(label: "Geographic data of buildings");

            int order = 0;

            // Instructions
            string intstr = "";
            intstr = "Fill in the building specific data below. ";
            intstr += "Use the checkboxes to indicate what types of renovation procedures you want to perform for this alternative. ";
            intstr += "You need to fill in the  building properties as well as the parameters under checked checkboxes. ";
            intstr += "If this is the as-is step leave all checkboxes unchecked and fill in only the building properties. ";
            intstr += "If multiple buildings have common properties you may select all of them and assign them values simultaneously. ";
            InputGroup instructions = new InputGroup(label: intstr, order: ++order);
            buildning_specific_data.Add(key: "instructions", item: instructions);

            // Building Common
            ++order;
            BuildingProperties(ref buildning_specific_data, ref order);

            // Heating System
            ++order;
            HeatingSystem(ref buildning_specific_data, ref order);

            // Building Shell
            ++order;
            BuildingShell(ref buildning_specific_data, ref order);

            // Ventilation System
            ++order;
            VentilationSystem(ref buildning_specific_data, ref order);

            //
            ++order;
            buildning_specific_data.Add(key: "changes_ian_ahe_and_aed", item: new InputGroup(label: "Changes due to renovation of building shell, ventilation and/or circulation pump.", order: ++order));
            buildning_specific_data.Add(key: change_in_ahd_due_to_renovations_of_bshell_ventilation_pump, item: new Number(label: change_in_ahd_due_to_renovations_of_bshell_ventilation_pump_lbl, unit: "kWh/year", order: ++order));
            buildning_specific_data.Add(key: change_in_aed_due_to_renovations_of_bshell_ventilation_pump, item: new Number(label: change_in_aed_due_to_renovations_of_bshell_ventilation_pump_lbl, unit: "MWh/year", order: ++order));


            // Radiators, pipes and electricity
            ++order;
            RadiatorsPipesElectricity(ref buildning_specific_data, ref order);

            return buildning_specific_data;
        }

        void BuildingProperties(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("building_properties", new InputGroup("Building Properties", order: ++order));

            // Inputs required in all cases
            input.Add(key: heat_source_before, item: new Select(label: heat_source_before_lbl, options: heat_sources, value: heat_sources.Last(), order: ++order));
            input.Add(key: heated_area, item: new Number(label: heated_area_lbl, min: 1, unit: "m\u00b2", order: ++order, value: 99));
            //input.Add(key: nr_apartments, item: new Number(label: nr_apartments_lbl, min: 1, order: ++order, value: 98));
                        
        }

        void HeatingSystem(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("heating_system", new InputGroup("Renovate Heating System", ++order));

            //input.Add(key: "1", item: new InputGroup(label: "-      Change Heating System", order: ++order));

            // Change Heating System
            input.Add(key: change_heating_system, item: new Checkbox(label: change_heating_system_lbl, order: ++order));
            input.Add(key: heat_source_after, item: new Select(label: heat_source_after_lbl, options: heat_sources, value: heat_sources.First(), order: ++order));
            input.Add(key: heating_system_life_of_product, item: new Number(label: heating_system_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: ahd_after_renovation, item: new Number(label: ahd_after_renovation_lbl, min: 0, unit: "kWh/year", order: ++order));
            input.Add(key: design_capacity, item: new Number(label: design_capacity_lbl, min: 0, unit: "kW", order: ++order));
            input.Add(key: weight_of_bhd, item: new Number(label: weight_of_bhd_lbl, min: 0, unit: "kg", order: ++order));
            input.Add(key: depth_of_borehole, item: new Number(label: depth_of_borehole_lbl, min: 0, unit: "m", order: ++order));
            //input.Add(key: heating_system_transport_to_building_truck, item: new Number(label: heating_system_transport_to_building_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: heating_system_transport_to_building_train, item: new Number(label: heating_system_transport_to_building_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: heating_system_transport_to_building_ferry, item: new Number(label: heating_system_transport_to_building_ferry_lbl, min: 0, unit: "km", order: ++order));

            //input.Add(key: "2", item: new InputGroup(label: "-      Change Circulation Pump", order: ++order));

            // Change Circulation Pump
            input.Add(key: change_circulationpump_in_heating_system, item: new Checkbox(label: change_circulationpump_in_heating_system_lbl, order: ++order));
            //input.Add(key: type_of_control_in_heating_system, item: new Select(label: type_of_control_in_heating_system_lbl, options: type_of_flow_control_in_heating_system_opts, order: ++order));
            input.Add(key: circulationpump_life_of_product, item: new Number(label: circulationpump_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: design_pressure_head, item: new Number(label: design_pressure_head_lbl, min: 0, unit: "kPa", order: ++order));
            //input.Add(key: design_flow_rate, item: new Number(label: design_flow_rate_lbl, min: 0, unit: "m\u00b3/h", order: ++order));
            input.Add(key: weight, item: new Number(label: weight_lbl, min: 0, unit: "kg", order: ++order));
            //input.Add(key: circulationpump_transport_to_building_truck, item: new Number(label: circulationpump_transport_to_building_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: circulationpump_transport_to_building_train, item: new Number(label: circulationpump_transport_to_building_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: circulationpump_transport_to_building_ferry, item: new Number(label: circulationpump_transport_to_building_ferry_lbl, min: 0, unit: "km", order: ++order));

        }

        void BuildingShell(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("building_shell", new InputGroup("Renovate Building Shell", ++order));                      

            // Insulation material 1
            input.Add(key: change_insulation_material_1, item: new Checkbox(label: change_insulation_material_1_lbl, order: ++order));
            input.Add(key: insulation_material_1_type_of_insulation, item: new Select(label: insulation_material_1_type_of_insulation_lbl, options: type_of_insulation, order: ++order));
            input.Add(key: insulation_material_1_life_of_product, item: new Number(label: insulation_material_1_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: insulation_material_1_change_in_annual_heat_demand_due_to_insulation, item: new Number(label: insulation_material_1_change_in_annual_heat_demand_due_to_insulation_lbl, unit: "kWh/year", order: ++order));
            input.Add(key: insulation_material_1_amount_of_new_insulation_material, item: new Number(label: insulation_material_1_amount_of_new_insulation_material_lbl, min: 0, unit: "kg", order: ++order));
            //input.Add(key: insulation_material_1_transport_to_building_by_truck, item: new Number(label: insulation_material_1_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: insulation_material_1_transport_to_building_by_train, item: new Number(label: insulation_material_1_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: insulation_material_1_transport_to_building_by_ferry, item: new Number(label: insulation_material_1_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Insulation material 2
            input.Add(key: change_insulation_material_2, item: new Checkbox(label: change_insulation_material_2_lbl, order: ++order));
            input.Add(key: insulation_material_2_type_of_insulation, item: new Select(label: insulation_material_2_type_of_insulation_lbl, options: type_of_insulation, order: ++order));
            input.Add(key: insulation_material_2_life_of_product, item: new Number(label: insulation_material_2_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: insulation_material_2_change_in_annual_heat_demand_due_to_insulation, item: new Number(label: insulation_material_2_change_in_annual_heat_demand_due_to_insulation_lbl, unit: "kWh/year", order: ++order));
            input.Add(key: insulation_material_2_amount_of_new_insulation_material, item: new Number(label: insulation_material_2_amount_of_new_insulation_material_lbl, min: 0, unit: "kg", order: ++order));
            //input.Add(key: insulation_material_2_transport_to_building_by_truck, item: new Number(label: insulation_material_2_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: insulation_material_2_transport_to_building_by_train, item: new Number(label: insulation_material_2_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: insulation_material_2_transport_to_building_by_ferry, item: new Number(label: insulation_material_2_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Fascade System
            input.Add(key: change_facade_system, item: new Checkbox(label: change_facade_system_lbl, order: ++order));
            input.Add(key: facade_system_type_facade_system, item: new Select(label: facade_system_type_of_facade_system_lbl, options: type_of_facade_system, order: ++order));
            input.Add(key: facade_system_life_of_product, item: new Number(label: facade_system_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: facade_system_change_in_annual_heat_demand_due_to_facade_system, item: new Number(label: facade_system_change_in_annual_heat_demand_due_to_facade_system_lbl, unit: "kWh/year", order: ++order));
            input.Add(key: facade_system_area_of_new_facade_system, item: new Number(label: facade_system_area_of_new_facade_system_lbl, min: 0, unit: "m\u00b2", order: ++order));
            //input.Add(key: facade_system_transport_to_building_by_truck, item: new Number(label: facade_system_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: facade_system_transport_to_building_by_train, item: new Number(label: facade_system_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: facade_system_transport_to_building_by_ferry, item: new Number(label: facade_system_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Windows
            input.Add(key: change_windows, item: new Checkbox(label: change_windows_lbl, order: ++order));
            input.Add(key: windows_type_windows, item: new Select(label: windows_type_of_windows_lbl, options: type_of_windows, order: ++order));
            input.Add(key: windows_life_of_product, item: new Number(label: windows_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: windows_change_in_annual_heat_demand_due_to_windows, item: new Number(label: windows_change_in_annual_heat_demand_due_to_windows_lbl, unit: "kWh/year", order: ++order));
            input.Add(key: windows_area_of_new_windows, item: new Number(label: windows_area_of_new_windows_lbl, min: 0, unit: "m\u00b2", order: ++order));
            //input.Add(key: windows_transport_to_building_by_truck, item: new Number(label: windows_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: windows_transport_to_building_by_train, item: new Number(label: windows_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: windows_transport_to_building_by_ferry, item: new Number(label: windows_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Doors
            input.Add(key: change_doors, item: new Checkbox(label: change_doors_lbl, order: ++order));
            input.Add(key: doors_type_doors, item: new Select(label: doors_type_of_doors_lbl, options: type_of_doors, order: ++order));
            input.Add(key: doors_life_of_product, item: new Number(label: doors_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: doors_change_in_annual_heat_demand_due_to_doors, item: new Number(label: doors_change_in_annual_heat_demand_due_to_doors_lbl, unit: "kWh/year", order: ++order));
            input.Add(key: doors_number_of_new_front_doors, item: new Number(label: doors_number_of_new_front_doors_lbl, min: 0, order: ++order));
            //input.Add(key: doors_transport_to_building_by_truck, item: new Number(label: doors_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: doors_transport_to_building_by_train, item: new Number(label: doors_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: doors_transport_to_building_by_ferry, item: new Number(label: doors_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

        }

        void VentilationSystem(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("ventilation_system", new InputGroup("Renovate Ventilation System", ++order));
            
            //Ventilation renovation
            //input.Add(key: ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation, item: new Number(label: ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation_lbl, min: 0, unit: "kWh/year", order: ++order));
            //input.Add(key: ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation, item: new Number(label: ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation_lbl, min: 0, unit: "kWh/year", order: ++order));

            // Ventilation ducts
            input.Add(key: change_ventilation_ducts, item: new Checkbox(label: change_ventilation_ducts_lbl, order: ++order));
            input.Add(key: ventilation_ducts_type_of_material, item: new Select(label: ventilation_ducts_type_of_material_lbl, options: type_of_ventilation_ducts_material, order: ++order));
            input.Add(key: ventilation_ducts_life_of_product, item: new Number(label: ventilation_ducts_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: ventilation_ducts_weight_of_ventilation_ducts, item: new Number(label: ventilation_ducts_weight_of_ventilation_ducts_lbl, unit: "kWh/year", order: ++order));
            //input.Add(key: ventilation_ducts_transport_to_building_by_truck, item: new Number(label: ventilation_ducts_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: ventilation_ducts_transport_to_building_by_train, item: new Number(label: ventilation_ducts_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: ventilation_ducts_transport_to_building_by_ferry, item: new Number(label: ventilation_ducts_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Airflow assembly
            input.Add(key: change_airflow_assembly, item: new Checkbox(label: change_airflow_assembly_lbl, order: ++order));
            input.Add(key: airflow_assembly_type_of_airflow_assembly, item: new Select(label: airflow_assembly_type_of_airflow_assembly_lbl, options: type_of_airflow_assembly, order: ++order));
            input.Add(key: airflow_assembly_life_of_product, item: new Number(label: airflow_assembly_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: airflow_assembly_design_airflow_exhaust_air, item: new Number(label: airflow_assembly_design_airflow_exhaust_air_lbl, unit: "kWh/year", order: ++order));
            //input.Add(key: airflow_assembly_transport_to_building_by_truck, item: new Number(label: airflow_assembly_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: airflow_assembly_transport_to_building_by_train, item: new Number(label: airflow_assembly_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: airflow_assembly_transport_to_building_by_ferry, item: new Number(label: airflow_assembly_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Air distribution housings and silencers
            input.Add(key: change_air_distribution_housings_and_silencers, item: new Checkbox(label: change_air_distribution_housings_and_silencers_lbl, order: ++order));            
            input.Add(key: air_distribution_housings_and_silencers_number_of_distribution_housings, item: new Number(label: air_distribution_housings_and_silencers_number_of_distribution_housings_lbl, min: 0, order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_life_of_product, item: new Number(label: air_distribution_housings_and_silencers_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: air_distribution_housings_and_silencers_transport_to_building_by_truck, item: new Number(label: air_distribution_housings_and_silencers_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: air_distribution_housings_and_silencers_transport_to_building_by_train, item: new Number(label: air_distribution_housings_and_silencers_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: air_distribution_housings_and_silencers_transport_to_building_by_ferry, item: new Number(label: air_distribution_housings_and_silencers_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));
            
        }

        void RadiatorsPipesElectricity(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("radiators_pipes_and_electricity", new InputGroup("Renovate Radiators, Pipes and/or Electricity", ++order));

            // Radiators
            input.Add(key: change_radiators, item: new Checkbox(label: change_radiators_lbl, order: ++order));
            input.Add(key: radiators_type_of_radiators, item: new Select(label: radiators_type_of_radiators_lbl, options: type_of_radiators, order: ++order));
            input.Add(key: radiators_life_of_product, item: new Number(label: radiators_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: radiators_weight_of_radiators, item: new Number(label: radiators_weight_of_radiators_lbl, unit: "kg", order: ++order));
            //input.Add(key: radiators_transport_to_building_by_truck, item: new Number(label: radiators_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: radiators_transport_to_building_by_train, item: new Number(label: radiators_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: radiators_transport_to_building_by_ferry, item: new Number(label: radiators_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - Copper
            input.Add(key: change_piping_copper, item: new Checkbox(label: change_piping_copper_lbl, order: ++order));
            input.Add(key: piping_copper_life_of_product, item: new Number(label: piping_copper_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: piping_copper_weight_of_copper_pipes, item: new Number(label: piping_copper_weight_of_copper_pipes_lbl, unit: "kg", order: ++order));
            //input.Add(key: piping_copper_transport_to_building_by_truck, item: new Number(label: piping_copper_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_copper_transport_to_building_by_train, item: new Number(label: piping_copper_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_copper_transport_to_building_by_ferry, item: new Number(label: piping_copper_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - PEX
            input.Add(key: change_piping_pex, item: new Checkbox(label: change_piping_pex_lbl, order: ++order));
            input.Add(key: piping_pex_life_of_product, item: new Number(label: piping_pex_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: piping_pex_weight_of_pex_pipes, item: new Number(label: piping_pex_weight_of_pex_pipes_lbl, unit: "kg", order: ++order));
            //input.Add(key: piping_pex_transport_to_building_by_truck, item: new Number(label: piping_pex_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_pex_transport_to_building_by_train, item: new Number(label: piping_pex_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_pex_transport_to_building_by_ferry, item: new Number(label: piping_pex_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - PP
            input.Add(key: change_piping_pp, item: new Checkbox(label: change_piping_pp_lbl, order: ++order));
            input.Add(key: piping_pp_life_of_product, item: new Number(label: piping_pp_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: piping_pp_weight_of_pp_pipes, item: new Number(label: piping_pp_weight_of_pp_pipes_lbl, unit: "kg", order: ++order));
            //input.Add(key: piping_pp_transport_to_building_by_truck, item: new Number(label: piping_pp_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_pp_transport_to_building_by_train, item: new Number(label: piping_pp_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_pp_transport_to_building_by_ferry, item: new Number(label: piping_pp_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - Cast Iron
            input.Add(key: change_piping_cast_iron, item: new Checkbox(label: change_piping_cast_iron_lbl, order: ++order));
            input.Add(key: piping_cast_iron_life_of_product, item: new Number(label: piping_cast_iron_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: piping_cast_iron_weight_of_cast_iron_pipes, item: new Number(label: piping_cast_iron_weight_of_cast_iron_pipes_lbl, unit: "kg", order: ++order));
            //input.Add(key: piping_cast_iron_transport_to_building_by_truck, item: new Number(label: piping_cast_iron_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_cast_iron_transport_to_building_by_train, item: new Number(label: piping_cast_iron_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_cast_iron_transport_to_building_by_ferry, item: new Number(label: piping_cast_iron_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - Galvanized Steel
            input.Add(key: change_piping_galvanized_steel, item: new Checkbox(label: change_piping_galvanized_steel_lbl, order: ++order));
            input.Add(key: piping_galvanized_steel_life_of_product, item: new Number(label: piping_galvanized_steel_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: piping_galvanized_steel_weight_of_galvanized_steel_pipes, item: new Number(label: piping_galvanized_steel_weight_of_galvanized_steel_pipes_lbl, unit: "kg", order: ++order));
            //input.Add(key: piping_galvanized_steel_transport_to_building_by_truck, item: new Number(label: piping_galvanized_steel_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_galvanized_steel_transport_to_building_by_train, item: new Number(label: piping_galvanized_steel_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_galvanized_steel_transport_to_building_by_ferry, item: new Number(label: piping_galvanized_steel_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Piping System - Relining
            input.Add(key: change_piping_relining, item: new Checkbox(label: change_piping_relining_lbl, order: ++order));
            input.Add(key: piping_relining_life_of_product, item: new Number(label: piping_relining_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: piping_relining_weight_of_relining_pipes, item: new Number(label: piping_relining_weight_of_relining_pipes_lbl, unit: "kg", order: ++order));
            //input.Add(key: piping_relining_transport_to_building_by_truck, item: new Number(label: piping_relining_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_relining_transport_to_building_by_train, item: new Number(label: piping_relining_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: piping_relining_transport_to_building_by_ferry, item: new Number(label: piping_relining_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

            // Electrical wiring
            input.Add(key: change_electrical_wiring, item: new Checkbox(label: change_electrical_wiring_lbl, order: ++order));
            input.Add(key: electrical_wiring_life_of_product, item: new Number(label: electrical_wiring_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: electrical_wiring_weight_of_electrical_wiring, item: new Number(label: electrical_wiring_weight_of_electrical_wiring_lbl, unit: "kg", order: ++order));
            //input.Add(key: electrical_wiring_transport_to_building_by_truck, item: new Number(label: electrical_wiring_transport_to_building_by_truck_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: electrical_wiring_transport_to_building_by_train, item: new Number(label: electrical_wiring_transport_to_building_by_train_lbl, min: 0, unit: "km", order: ++order));
            //input.Add(key: electrical_wiring_transport_to_building_by_ferry, item: new Number(label: electrical_wiring_transport_to_building_by_ferry_lbl, min: 0, unit: "km", order: ++order));

        }

        void SetInputDataOneBuilding(Feature building, ref CExcel exls)
        {
            SetBuildingProperties(building, ref exls);
            SetHeatingSystem(building, ref exls);
            SetBuildingShell(building, ref exls);
            SetVentilationSystem(building, ref exls);

            #region Change...
            String Key;
            object value;

            Key = change_in_ahd_due_to_renovations_of_bshell_ventilation_pump;
            value = Convert.ToDouble(building.properties[Key]);
            Set(sheet: "Indata", cell: "C288", value: value, exls: ref exls);

            Key = change_in_aed_due_to_renovations_of_bshell_ventilation_pump;
            value = Convert.ToDouble(building.properties[Key]);
            Set(sheet: "Indata", cell: "C289", value: value, exls: ref exls);
            #endregion

            SetRadiatorsPipesElectricity(building, ref exls);
        }

        void SetBuildingProperties(Feature building, ref CExcel exls)
        {
            String Key;
            object value;
            String cell;

            // Inputs required in all cases
            #region Heated Area
            Key = heated_area;
            value = Convert.ToDouble(building.properties[Key]);
            cell = "C25";
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            //#region Number of Apartments
            //Key = nr_apartments;
            //cell = "C26";
            //value = Convert.ToDouble(building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            #region Heat Source Before
            Key = heat_source_before;
            cell = "C93";
            value = heat_sources.GetIndex((string)building.properties[Key]) + 1;
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion
            
        }

        void SetHeatingSystem(Feature building, ref CExcel exls)
        {

            String Key;
            object value;
            String cell;

            // Change Heating System
            #region Change Heating System
            Key = change_heating_system;
            cell = "C99";
            value = (bool)building.properties[Key];
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion
            
            #region Heat Source After
            Key = heat_source_after;
            cell = "C94";
            value = heat_sources.GetIndex((string)building.properties[Key]) + 1;
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            #region Annual Heat Demand After Renovation
            Key = ahd_after_renovation;
            cell = "C95";
            value = Convert.ToDouble(building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            #region Heating System: Life of Product
            Key = heating_system_life_of_product;
            cell = "C100";
            value = Convert.ToDouble(building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            #region Design Capacity
            Key = design_capacity;
            cell = "C103";
            value = Convert.ToDouble(building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            #region Weight of boiler/heat pump/district heating substation
            Key = weight_of_bhd;
            cell = "C104";
            value = Convert.ToDouble(building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            #region For geothermal heat pump: Depth of bore hole
            Key = depth_of_borehole;
            cell = "C109";
            value = Convert.ToDouble(building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            //#region Transport to building by truck
            //Key = heating_system_transport_to_building_truck;
            //cell = "C106";
            //value = Convert.ToDouble(building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Transport to building by train
            //Key = heating_system_transport_to_building_train;
            //cell = "C107";
            //value = Convert.ToDouble(building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Transport to building by ferry
            //Key = heating_system_transport_to_building_ferry;
            //cell = "C108";
            //value = Convert.ToDouble(building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            // Change Circulation Pump
            #region Change Circulation Pump
            Key = change_circulationpump_in_heating_system;
            cell = "C113";
            value = (bool)(building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            #region Circulation Pump: Life of Product
            Key = circulationpump_life_of_product;
            cell = "C114";
            value = (building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            //#region Design pressure head
            //Key = design_pressure_head;
            //cell = "C115";
            //value = (building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Design flow rate
            //Key = design_flow_rate;
            //cell = "C116";
            //value = (building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Type of flow control in heating system
            //Key = type_of_control_in_heating_system;
            //cell = "C117";
            //value = type_of_flow_control_in_heating_system_opts.GetIndex((string)building.properties[Key]) + 1;
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            #region Weight
            Key = weight;
            cell = "C118";
            value = (building.properties[Key]);
            if (!exls.SetCellValue("Indata", cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            #endregion

            //#region Transport to building by truck
            //Key = circulationpump_transport_to_building_truck;
            //cell = "C120";
            //value = (building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Transport to building by train
            //Key = circulationpump_transport_to_building_train;
            //cell = "C121";
            //value = (building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Transport to building by ferry
            //Key = circulationpump_transport_to_building_ferry;
            //cell = "C122";
            //value = (building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

        }

        void SetBuildingShell(Feature building, ref CExcel exls)
        {

            String Key;
            object value;
            String cell;

            // Insulation material 1
            #region Change Insulation Material 1
            #region Change Insulation Material 1?
            Key = change_insulation_material_1;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C126", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Change Insulation Material 1: Life of Product
                Key = insulation_material_1_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C127", value: value, exls: ref exls);
                #endregion

                #region Change Insulation Material 1: Type of Material
                Key = insulation_material_1_type_of_insulation;
                value = type_of_insulation.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C128", value: value, exls: ref exls);
                #endregion

                //#region Change Insulation Material 1: Change AHD due to New Insulation
                //Key = insulation_material_1_change_in_annual_heat_demand_due_to_insulation;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C129", value: value, exls: ref exls);
                //#endregion

                #region Change Insulation Material 1: Amount of Insulation Material
                Key = insulation_material_1_amount_of_new_insulation_material;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C130", value: value, exls: ref exls);
                #endregion

                //#region Change Insulation Material 1: Transport by Truck [km]
                //Key = insulation_material_1_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C132", value: value, exls: ref exls);
                //#endregion

                //#region Change Insulation Material 1: Transport by Truck [km]
                //Key = insulation_material_1_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C133", value: value, exls: ref exls);
                //#endregion

                //#region Change Insulation Material 1: Transport by Truck [km]
                //Key = insulation_material_1_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C134", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Insulation material 2
            #region Change Insulation Material 2
            #region Change Insulation Material 2?
            Key = change_insulation_material_2;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C137", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Insulation Material 2: Life of Product
                Key = insulation_material_2_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C138", value: value, exls: ref exls);
                #endregion

                #region Insulation Material 2: Type of Material
                Key = insulation_material_2_type_of_insulation;
                value = type_of_insulation.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C139", value: value, exls: ref exls);
                #endregion

                //#region Insulation Material 2: Change AHD due to New Insulation
                //Key = insulation_material_2_change_in_annual_heat_demand_due_to_insulation;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C140", value: value, exls: ref exls);
                //#endregion

                #region Insulation Material 2: Amount of Insulation Material
                Key = insulation_material_2_amount_of_new_insulation_material;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C141", value: value, exls: ref exls);
                #endregion

                //#region Insulation Material 2: Transport by Truck [km]
                //Key = insulation_material_2_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C143", value: value, exls: ref exls);
                //#endregion

                //#region Insulation Material 2: Transport by Truck [km]
                //Key = insulation_material_2_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C144", value: value, exls: ref exls);
                //#endregion

                //#region Insulation Material 2: Transport by Truck [km]
                //Key = insulation_material_2_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C145", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Facade System
            #region Change Fascade System
            #region Change Fascade System?
            Key = change_facade_system;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C148", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Fascade System: Life of Product
                Key = facade_system_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C149", value: value, exls: ref exls);
                #endregion

                #region Fascade System: Type of Fascade System
                Key = facade_system_type_facade_system;
                value = type_of_facade_system.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C150", value: value, exls: ref exls);
                #endregion

                //#region Fascade System: Change AHD due to New Fascade System
                //Key = facade_system_change_in_annual_heat_demand_due_to_facade_system;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C151", value: value, exls: ref exls);
                //#endregion

                #region Fascade System: Area of New Fascade System
                Key = facade_system_area_of_new_facade_system;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C152", value: value, exls: ref exls);
                #endregion

                //#region Fascade System: Transport by Truck [km]
                //Key = facade_system_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C154", value: value, exls: ref exls);
                //#endregion

                //#region Fascade System: Transport by Truck [km]
                //Key = facade_system_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C155", value: value, exls: ref exls);
                //#endregion

                //#region Fascade System: Transport by Truck [km]
                //Key = facade_system_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C156", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Windows
            #region Change Windows
            #region Change Windows?
            Key = change_windows;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C159", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Windows: Life of Product
                Key = windows_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C160", value: value, exls: ref exls);
                #endregion

                #region Windows: Type of Windows
                Key = windows_type_windows;
                value = type_of_windows.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C161", value: value, exls: ref exls);
                #endregion

                //#region Windows: Change AHD due to New Windows
                //Key = windows_change_in_annual_heat_demand_due_to_windows;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C162", value: value, exls: ref exls);
                //#endregion

                #region Windows: Area of New Windows
                Key = windows_area_of_new_windows;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C163", value: value, exls: ref exls);
                #endregion

                //#region Windows: Transport by Truck [km]
                //Key = windows_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C165", value: value, exls: ref exls);
                //#endregion

                //#region Windows: Transport by Truck [km]
                //Key = windows_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C166", value: value, exls: ref exls);
                //#endregion

                //#region Windows: Transport by Truck [km]
                //Key = windows_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C167", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Doors
            #region Change Doors
            #region Change Doors?
            Key = change_doors;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C170", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Doors: Life of Product
                Key = doors_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C171", value: value, exls: ref exls);
                #endregion

                #region Doors: Type of Doors
                Key = doors_type_doors;
                value = type_of_doors.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C172", value: value, exls: ref exls);
                #endregion

                //#region Doors: Change AHD due to New Doors
                //Key = doors_change_in_annual_heat_demand_due_to_doors;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C173", value: value, exls: ref exls);
                //#endregion

                #region Doors: Number of new Fron Doors
                Key = doors_number_of_new_front_doors;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C174", value: value, exls: ref exls);
                #endregion

                //#region Doors: Transport by Truck [km]
                //Key = doors_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C176", value: value, exls: ref exls);
                //#endregion

                //#region Doors: Transport by Truck [km]
                //Key = doors_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C177", value: value, exls: ref exls);
                //#endregion

                //#region Doors: Transport by Truck [km]
                //Key = doors_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C178", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

        }

        void SetVentilationSystem(Feature building, ref CExcel exls)
        {

            String Key;
            object value;

            // - Ventilation System
            #region Ventilation System

            //#region Ventilation System: Change in AHD due to ventilation system renovation
            //Key = ventilation_change_in_annual_heat_demand_due_ventilation_systems_renovation;
            //value = Convert.ToDouble(building.properties[Key]);
            //Set(sheet: "Indata", cell: "C210", value: value, exls: ref exls);
            //#endregion

            //#region Ventilation System: Change in AED due to ventilation system renovation
            //Key = ventilation_change_in_annual_electricity_demand_due_ventilation_systems_renovation;
            //value = Convert.ToDouble(building.properties[Key]);
            //Set(sheet: "Indata", cell: "C211", value: value, exls: ref exls);
            //#endregion

            #endregion

            // Ventilation Ducts
            #region Ventilation Ducts
            #region Ventilation Ducts?
            Key = change_ventilation_ducts;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C182", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Ventilation Ducts: Life of Product
                Key = ventilation_ducts_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C183", value: value, exls: ref exls);
                #endregion

                #region Ventilation Ducts: Type of Material
                Key = ventilation_ducts_type_of_material;
                value = type_of_ventilation_ducts_material.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C184", value: value, exls: ref exls);
                #endregion

                #region Ventilation Ducts: Weight of Ventilation Ducts
                Key = ventilation_ducts_weight_of_ventilation_ducts;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C185", value: value, exls: ref exls);
                #endregion

                //#region Ventilation Ducts: Transport by Truck [km]
                //Key = ventilation_ducts_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C187", value: value, exls: ref exls);
                //#endregion

                //#region Ventilation Ducts: Transport by Truck [km]
                //Key = ventilation_ducts_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C188", value: value, exls: ref exls);
                //#endregion

                //#region Ventilation Ducts: Transport by Truck [km]
                //Key = ventilation_ducts_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C189", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Airflow Assembly
            #region Change Airflow Assembly
            #region Change Airflow Assembly?
            Key = change_airflow_assembly;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C192", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Airflow Assembly: Life of Product
                Key = airflow_assembly_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C193", value: value, exls: ref exls);
                #endregion

                #region Airflow Assembly: Type of Airflow Assembly
                Key = airflow_assembly_type_of_airflow_assembly;
                value = type_of_airflow_assembly.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C194", value: value, exls: ref exls);
                #endregion

                #region Airflow Assembly: Area of New Airflow Assembly
                Key = airflow_assembly_design_airflow_exhaust_air;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C195", value: value, exls: ref exls);
                #endregion

                //#region Airflow Assembly: Transport by Truck [km]
                //Key = airflow_assembly_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C197", value: value, exls: ref exls);
                //#endregion

                //#region Airflow Assembly: Transport by Truck [km]
                //Key = airflow_assembly_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C198", value: value, exls: ref exls);
                //#endregion

                //#region Airflow Assembly: Transport by Truck [km]
                //Key = airflow_assembly_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C199", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Air Distribution Housings & Silencers
            #region Change Air Distribution Housings & Silencers
            #region Change Air Distribution Housings & Silencers?
            Key = change_air_distribution_housings_and_silencers;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C202", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Air Distribution Housings & Silencers: Number of Housings
                Key = air_distribution_housings_and_silencers_number_of_distribution_housings;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C208", value: value, exls: ref exls);  
                #endregion

                #region Air Distribution Housings & Silencers: Life of Product
                Key = air_distribution_housings_and_silencers_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C203", value: value, exls: ref exls);
                #endregion

                //#region Air Distribution Housings & Silencers: Transport by Truck [km]
                //Key = air_distribution_housings_and_silencers_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C205", value: value, exls: ref exls);
                //#endregion

                //#region Air Distribution Housings & Silencers: Transport by Truck [km]
                //Key = air_distribution_housings_and_silencers_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C206", value: value, exls: ref exls);
                //#endregion

                //#region Air Distribution Housings & Silencers: Transport by Truck [km]
                //Key = air_distribution_housings_and_silencers_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C207", value: value, exls: ref exls);
                //#endregion
            }
            #endregion
        }

        void SetRadiatorsPipesElectricity(Feature building, ref CExcel exls)
        {

            String Key;
            object value;

            // Radiators
            #region Radiators
            #region Radiators?
            Key = change_radiators;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C215", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Radiators: Life of Product
                Key = radiators_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C216", value: value, exls: ref exls);
                #endregion

                #region Radiators: Type of Material
                Key = radiators_type_of_radiators;
                value = type_of_radiators.GetIndex((string)building.properties[Key]) + 1;
                Set(sheet: "Indata", cell: "C217", value: value, exls: ref exls);
                #endregion

                #region Radiators: Weight of Radiators
                Key = radiators_weight_of_radiators;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C218", value: value, exls: ref exls);
                #endregion

                //#region Radiators: Transport by Truck [km]
                //Key = radiators_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C220", value: value, exls: ref exls);
                //#endregion

                //#region Radiators: Transport by Truck [km]
                //Key = radiators_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C221", value: value, exls: ref exls);
                //#endregion

                //#region Radiators: Transport by Truck [km]
                //Key = radiators_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C222", value: value, exls: ref exls);
                //#endregion
            }
            #endregion
            
            // Piping System Copper
            #region Change Piping System Copper
            #region Change Piping System Copper?
            Key = change_piping_copper;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C225", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Piping System Copper: Life of Product
                Key = piping_copper_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C226", value: value, exls: ref exls);
                #endregion

                #region Piping System Copper: Area of New Piping System Copper
                Key = piping_copper_weight_of_copper_pipes;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C227", value: value, exls: ref exls);
                #endregion

                //#region Piping System Copper: Transport by Truck [km]
                //Key = piping_copper_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C229", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Copper: Transport by Truck [km]
                //Key = piping_copper_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C230", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Copper: Transport by Truck [km]
                //Key = piping_copper_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C231", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Piping System PEX
            #region Change Piping System PEX
            #region Change Piping System PEX?
            Key = change_piping_pex;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C234", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Piping System PEX: Life of Product
                Key = piping_pex_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C235", value: value, exls: ref exls);
                #endregion

                #region Piping System PEX: Area of New Piping System PEX
                Key = piping_pex_weight_of_pex_pipes;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C236", value: value, exls: ref exls);
                #endregion

                //#region Piping System PEX: Transport by Truck [km]
                //Key = piping_pex_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C238", value: value, exls: ref exls);
                //#endregion

                //#region Piping System PEX: Transport by Truck [km]
                //Key = piping_pex_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C239", value: value, exls: ref exls);
                //#endregion

                //#region Piping System PEX: Transport by Truck [km]
                //Key = piping_pex_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C240", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Piping System PP
            #region Change Piping System PP
            #region Change Piping System PP?
            Key = change_piping_pp;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C243", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Piping System PP: Life of Product
                Key = piping_pp_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C244", value: value, exls: ref exls);
                #endregion

                #region Piping System PP: Area of New Piping System PP
                Key = piping_pp_weight_of_pp_pipes;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C245", value: value, exls: ref exls);
                #endregion

                //#region Piping System PP: Transport by Truck [km]
                //Key = piping_pp_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C247", value: value, exls: ref exls);
                //#endregion

                //#region Piping System PP: Transport by Truck [km]
                //Key = piping_pp_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C248", value: value, exls: ref exls);
                //#endregion

                //#region Piping System PP: Transport by Truck [km]
                //Key = piping_pp_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C249", value: value, exls: ref exls);
                //#endregion
            }
            #endregion
            
            // Piping System Cast Iron
            #region Change Piping System Cast Iron
            #region Change Piping System Cast Iron?
            Key = change_piping_cast_iron;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C252", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Piping System Cast Iron: Life of Product
                Key = piping_cast_iron_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C253", value: value, exls: ref exls);
                #endregion

                #region Piping System Cast Iron: Area of New Piping System Cast Iron
                Key = piping_cast_iron_weight_of_cast_iron_pipes;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C254", value: value, exls: ref exls);
                #endregion

                //#region Piping System Cast Iron: Transport by Truck [km]
                //Key = piping_cast_iron_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C256", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Cast Iron: Transport by Truck [km]
                //Key = piping_cast_iron_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C257", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Cast Iron: Transport by Truck [km]
                //Key = piping_cast_iron_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C258", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Piping System Galvanized Steel
            #region Change Piping System Galvanized Steel
            #region Change Piping System Galvanized Steel?
            Key = change_piping_galvanized_steel;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C261", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Piping System Galvanized Steel: Life of Product
                Key = piping_galvanized_steel_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C262", value: value, exls: ref exls);
                #endregion

                #region Piping System Galvanized Steel: Area of New Piping System Galvanized Steel
                Key = piping_galvanized_steel_weight_of_galvanized_steel_pipes;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C263", value: value, exls: ref exls);
                #endregion

                //#region Piping System Galvanized Steel: Transport by Truck [km]
                //Key = piping_galvanized_steel_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C265", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Galvanized Steel: Transport by Truck [km]
                //Key = piping_galvanized_steel_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C266", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Galvanized Steel: Transport by Truck [km]
                //Key = piping_galvanized_steel_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C267", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Piping System Relining
            #region Change Piping System Relining
            #region Change Piping System Relining?
            Key = change_piping_relining;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C270", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Piping System Relining: Life of Product
                Key = piping_relining_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C271", value: value, exls: ref exls);
                #endregion

                #region Piping System Relining: Area of New Piping System Relining
                Key = piping_relining_weight_of_relining_pipes;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C272", value: value, exls: ref exls);
                #endregion

                //#region Piping System Relining: Transport by Truck [km]
                //Key = piping_relining_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C274", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Relining: Transport by Truck [km]
                //Key = piping_relining_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C275", value: value, exls: ref exls);
                //#endregion

                //#region Piping System Relining: Transport by Truck [km]
                //Key = piping_relining_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C276", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

            // Electrical Wiring
            #region Change Electrical Wiring
            #region Change Electrical Wiring?
            Key = change_electrical_wiring;
            value = (bool)building.properties[Key];
            Set(sheet: "Indata", cell: "C279", value: value, exls: ref exls);
            #endregion
            if ((bool)value)
            {
                #region Electrical Wiring: Life of Product
                Key = electrical_wiring_life_of_product;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C280", value: value, exls: ref exls);
                #endregion

                #region Electrical Wiring: Area of New Electrical Wiring
                Key = electrical_wiring_weight_of_electrical_wiring;
                value = Convert.ToDouble(building.properties[Key]);
                Set(sheet: "Indata", cell: "C281", value: value, exls: ref exls);
                #endregion

                //#region Electrical Wiring: Transport by Truck [km]
                //Key = electrical_wiring_transport_to_building_by_truck;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C283", value: value, exls: ref exls);
                //#endregion

                //#region Electrical Wiring: Transport by Truck [km]
                //Key = electrical_wiring_transport_to_building_by_train;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C284", value: value, exls: ref exls);
                //#endregion

                //#region Electrical Wiring: Transport by Truck [km]
                //Key = electrical_wiring_transport_to_building_by_ferry;
                //value = Convert.ToDouble(building.properties[Key]);
                //Set(sheet: "Indata", cell: "C285", value: value, exls: ref exls);
                //#endregion
            }
            #endregion

        }

        private void Set(string sheet, string cell, object value, ref CExcel exls)
        {
            if (!exls.SetCellValue(sheet, cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {2} in sheet {3}", cell, value, sheet));
        }

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            if (!inputSpecifications.ContainsKey(kpiId))
                throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.", kpiId));

            return inputSpecifications[kpiId];
        }

        protected override Outputs CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls)
        {
            Outputs outputs = new Outputs();

            InputGroup commonPropertiesIpg = indata[common_properties] as InputGroup;
            Dictionary<String, Input> commonProperties = commonPropertiesIpg.GetInputs();
            GeoJson buildingProperties = indata["buildings"] as GeoJson;

            double kpi = 0;
            string resultCell;

            switch (kpiId)
            {
                case kpi_gwp:
                    resultCell = "C31"; //Change of global warming potential
                    break;
                case kpi_peu:
                    resultCell = "C32"; //Change of primary energy use  
                    break;
                default:
                    throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
            }

            #region Set Common Properties
            String Key;
            object value = 0;

            #region LCA Calculation Period
            Key = lca_calculation_period;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "Indata", cell: "C16", value: value, exls: ref exls);
            #endregion

            #region Electricity Mix
            Key = electricity_mix;
            value = ((Select)commonProperties[Key]).SelectedIndex() + 1;
            Set(sheet: "Indata", cell: "C17", value: value, exls: ref exls);
            #endregion

            // If district heating is used (before/after renovation)
            #region Global warming potential of district heating
            Key = gwp_district;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "Indata", cell: "C20", value: value, exls: ref exls);
            #endregion

            #region "Primary energy use of district heating
            Key = peu_district;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "Indata", cell: "C21", value: value, exls: ref exls);
            #endregion
            #endregion

            foreach (Feature building in buildingProperties.value.features)
            {
                if ((bool)building.properties[change_heating_system] ||
                    (bool)building.properties[change_circulationpump_in_heating_system] ||
                    (bool)building.properties[change_insulation_material_1] ||
                    (bool)building.properties[change_insulation_material_2] ||
                    (bool)building.properties[change_facade_system] ||
                    (bool)building.properties[change_windows] ||
                    (bool)building.properties[change_doors] ||
                    (bool)building.properties[change_ventilation_ducts] ||
                    (bool)building.properties[change_airflow_assembly] ||
                    (bool)building.properties[change_air_distribution_housings_and_silencers] ||
                    (bool)building.properties[change_radiators] ||
                    (bool)building.properties[change_piping_copper] ||
                    (bool)building.properties[change_piping_pex] ||
                    (bool)building.properties[change_piping_pp] ||
                    (bool)building.properties[change_piping_cast_iron] ||
                    (bool)building.properties[change_piping_galvanized_steel] ||
                    (bool)building.properties[change_piping_relining] ||
                    (bool)building.properties[change_electrical_wiring]) 
                {
                    SetInputDataOneBuilding(building, ref exls);

                    var resi = exls.GetCellValue("Indata", resultCell);
                    kpi += Convert.ToDouble(resi);

                }

            }

            switch (kpiId)
            {
                case kpi_gwp:
                    outputs.Add(new Kpi(Math.Round(kpi, 2), "Change of global warming potential", "tonnes CO2 eq"));
                    break;
                case kpi_peu:
                    outputs.Add(new Kpi(Math.Round(kpi, 2), "Change of primary energy use", "MWh"));
                    break;
                default:
                    throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Yaml.Serialization;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace MSR_LCC
{
    class LCC_Module : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_lcc = "lcc";
        const string result_cell = "E30";

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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
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
                LCC_Module_ErrorRaised(this, ex);
            }
        }

        void DefineInputSpecifications()
        {
            try
            {
                inputSpecifications = new Dictionary<string, InputSpecification>();

                //GeoJson
                inputSpecifications.Add(kpi_lcc, GetInputSpecificationGeoJson());
            }
            catch (System.Exception ex)
            {
                LCC_Module_ErrorRaised(this, ex);
            }
        }

        #region Input Specification (Names and Labels)
        // - Input Specification (Names and Labels)
        #region Common Properties
        // Common
        string common_properties = "common_properties";
        //string lcc_calculation_period = "lcc_calculation_period";
        string discount_rate = "discount_rate";
        string discount_rate_lbl = "Discount rate";
        string electric_cost = "elictric_cost";
        string electric_cost_lbl = "Electric cost";
        string water_cost = "water_cost";
        string water_cost_lbl = "Water cost";
        string heat_cost = "heat_cost";
        string heat_cost_lbl = "Heat cost";
        string natural_gas_cost = "natural_gas_cost";
        string natural_gas_cost_lbl = "Natural gas cost";

        #endregion

        #region Building specific Properties
        // Building specific Properties
        string buildings = "buildings";
        #region Building Common
        //Building Common
        // Inputs required in all cases
        //string heated_area = "HEATED_AREA";
        //string heated_area_lbl = "Heated area";
        //string heat_source_before = "HEAT_SOURCE_BEFORE";
        //string heat_source_before_lbl = "Heat source before renovation";
        //string heat_source_after = "HEAT_SOURCE_AFTER";  
        //string heat_source_after_lbl = "Heat source after renovation";
        // If district heating is used (before/after renovation)
        //string ep_district = "energy_price";
        //string ep_district_lbl = "Energy price";
        //string peu_district = "PRIMARY_ENERGY_USE_OF_DISTRICT_HEATING";
        //string peu_district_lbl = "Primary energy factor of district heating. Required if any building uses district heating before or after renovation. Impact per unit energy delivered to building, i.e. including distribution losses.";
        #endregion

        #region Heating System
        //Heating System
        // Change Heating System
        string change_heating_system = "HEATING_SYSTEM__CHANGE_HEATING_SYSTEM";
        string change_heating_system_lbl = "Replace building heating system";
        //string ahd_after_renovation = "HEATING_SYSTEM__AHD_AFTER_RENOVATION"; //TODO use instead of heat consuption and use?
        //string ahd_after_renovation_lbl = "Annual heat demand after renovation";
        string heating_system_life_of_product = "HEATING_SYSTEM__LIFE_OF_PRODUCT";
        string heating_system_life_of_product_lbl = "Life of product (Practical time of life of the products and materials used)";
        string heating_system_initial_investment = "heating_system_initial_investment";
        string heating_system_initial_investment_lbl = "Initial investment";
        string heating_system_installation_cost = "heating_system_installation_cost";
        string heating_system_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string heating_system_heat_consumption = "heating_system_heat_consumption";
        string heating_system_heat_consumption_lbl = "Heat consumption";
        string heating_system_heat_annual_use = "heating_system_heat_annual_use";
        string heating_system_heat_annual_use_lbl = "Annual use";  //Same for all components? (e.g. electric,water and natural gas)
        string heating_system_natural_gas_consumption = "heating_system_natural_gas_consumption";
        string heating_system_natural_gas_consumption_lbl = "Natural gas consumption";
        string heating_system_natural_gas_annual_use = "heating_system_natural_gas_annual_use";
        string heating_system_natural_gas_annual_use_lbl = "Natural gas annual use";
        string heating_system_electric_consumption = "heating_system_electric_consumption";
        string heating_system_electric_consumption_lbl = "Electric consumption";
        string heating_system_electric_annual_use = "heating_system_electric_annual_use";
        string heating_system_electric_annual_use_lbl = "Electric annual use";
        string heating_system_water_consumption = "heating_system_water_consumption";
        string heating_system_water_consumption_lbl = "Water consumption";
        string heating_system_water_annual_use = "heating_system_water_annual_use";
        string heating_system_water_annual_use_lbl = "Water annual use";
        string heating_system_maintenance_cost = "heating_system_maintenance_cost";
        string heating_system_maintenance_cost_lbl = "Total maintenance costs per year";
        string heating_system_taxes_fees_cost = "heating_system_taxes_fees_cost";
        string heating_system_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string heating_system_liquidation_cost = "heating_system_liquidation_cost";
        string heating_system_liquidation_cost_lbl = "Cost of liquidation";
        string heating_system_remnant_value = "heating_system_remnant_value";
        string heating_system_remnant_value_lbl = "Remnant value";

        // Change Circulation Pump
        string change_circulationpump_in_heating_system = "PUMP__CHANGE_PUMP";
        string change_circulationpump_in_heating_system_lbl = "Replace circulation pump in building heating system";
        string circulationpump_life_of_product = "PUMP__LIFE_OF_PRODUCT";
        string circulationpump_life_of_product_lbl = "Practical time of life of the products and materials used";
        string circulationpump_initial_investment = "circulationpump_initial_investment";
        string circulationpump_initial_investment_lbl = "Initial investment";
        string circulationpump_installation_cost = "circulationpump_installation_cost";
        string circulationpump_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string circulationpump_electric_consumption = "circulationpump_electric_consumption";
        string circulationpump_electric_consumption_lbl = "Electric consumption";
        string circulationpump_electric_annual_use = "circulationpump_electric_annual_use";
        string circulationpump_electric_annual_use_lbl = "Electric annual use";
        string circulationpump_maintenance_cost = "circulationpump_maintenance_cost";
        string circulationpump_maintenance_cost_lbl = "Total maintenance costs per year";
        string circulationpump_taxes_fees_cost = "circulationpump_taxes_fees_cost";
        string circulationpump_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string circulationpump_liquidation_cost = "circulationpump_liquidation_cost";
        string circulationpump_liquidation_cost_lbl = "Cost of liquidation";
        string circulationpump_remnant_value = "circulationpump_remnant_value";
        string circulationpump_remnant_value_lbl = "Remnant value";
        #endregion

        #region Building Shell
        //Building Shell
        // Insulation material 1
        string change_insulation_material_1 = "INSULATION_MATERIAL_ONE__CHANGE";
        string change_insulation_material_1_lbl = "Use insulation material 1";
        string insulation_material_1_life_of_product = "INSULATION_MATERIAL_ONE__LIFE_OF_PRODUCT";
        string insulation_material_1_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string insulation_material_1_type_of_insulation = "INSULATION_MATERIAL_ONE__TYPE_OF_INSULATION";
        //string insulation_material_1_type_of_insulation_lbl = "Type of insulation";
        //string insulation_material_1_amount_of_new_insulation_material = "INSULATION_MATERIAL_ONE__AMOUNT_OF_NEW_INSULATION_MATERIAL";
        //string insulation_material_1_amount_of_new_insulation_material_lbl = "Amount of new insulation material (required if renovation includes new insulation material)";
        string insulation_material_1_initial_investment = "insulation_material_1_initial_investment";
        string insulation_material_1_initial_investment_lbl = "Initial investment";
        string insulation_material_1_installation_cost = "insulation_material_1_installation_cost";
        string insulation_material_1_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string insulation_material_1_maintenance_cost = "insulation_material_1_maintenance_cost";
        string insulation_material_1_maintenance_cost_lbl = "Total maintenance costs per year";
        string insulation_material_1_taxes_fees_cost = "insulation_material_1_taxes_fees_cost";
        string insulation_material_1_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string insulation_material_1_liquidation_cost = "insulation_material_1_liquidation_cost";
        string insulation_material_1_liquidation_cost_lbl = "Cost of liquidation";
        string insulation_material_1_remnant_value = "insulation_material_1_remnant_value";
        string insulation_material_1_remnant_value_lbl = "Remnant value";

        // Insulation material 2
        string change_insulation_material_2 = "INSULATION_MATERIAL_TWO__CHANGE";
        string change_insulation_material_2_lbl = "Use insulation material 2";
        string insulation_material_2_life_of_product = "INSULATION_MATERIAL_TWO__LIFE_OF_PRODUCT";
        string insulation_material_2_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string insulation_material_2_type_of_insulation = "INSULATION_MATERIAL_ONE__TYPE_OF_INSULATION";
        //string insulation_material_2_type_of_insulation_lbl = "Type of insulation";
        //string insulation_material_2_amount_of_new_insulation_material = "INSULATION_MATERIAL_ONE__AMOUNT_OF_NEW_INSULATION_MATERIAL";
        //string insulation_material_2_amount_of_new_insulation_material_lbl = "Amount of new insulation material (required if renovation includes new insulation material)";
        string insulation_material_2_initial_investment = "insulation_material_2_initial_investment";
        string insulation_material_2_initial_investment_lbl = "Initial investment";
        string insulation_material_2_installation_cost = "insulation_material_2_installation_cost";
        string insulation_material_2_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string insulation_material_2_maintenance_cost = "insulation_material_2_maintenance_cost";
        string insulation_material_2_maintenance_cost_lbl = "Total maintenance costs per year";
        string insulation_material_2_taxes_fees_cost = "insulation_material_2_taxes_fees_cost";
        string insulation_material_2_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string insulation_material_2_liquidation_cost = "insulation_material_2_liquidation_cost";
        string insulation_material_2_liquidation_cost_lbl = "Cost of liquidation";
        string insulation_material_2_remnant_value = "insulation_material_2_remnant_value";
        string insulation_material_2_remnant_value_lbl = "Remnant value";

        // facade system
        string change_facade_system = "FACADE__CHANGE";
        string change_facade_system_lbl = "Change facade";
        string facade_system_life_of_product = "FACADE__LIFE_OF_PRODUCT";
        string facade_system_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string facade_system_type_facade_system = "FACADE__TYPE_OF_FACADE_SYSTEM";
        //string facade_system_type_of_facade_system_lbl = "Type of facade system";
        //string facade_system_change_in_annual_heat_demand_due_to_facade_system = "FACADE__CHANGE_IN_AHD_DUE_TO_FACADE_SYSTEM";
        //string facade_system_change_in_annual_heat_demand_due_to_facade_system_lbl = "Change in annual heat demand due to facade system (an energy saving is given as a negative value)";
        string facade_system_initial_investment = "facade_system_initial_investment";
        string facade_system_initial_investment_lbl = "Initial investment";
        string facade_system_installation_cost = "facade_system_installation_cost";
        string facade_system_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string facade_system_maintenance_cost = "facade_system_maintenance_cost";
        string facade_system_maintenance_cost_lbl = "Total maintenance costs per year";
        string facade_system_taxes_fees_cost = "facade_system_taxes_fees_cost";
        string facade_system_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string facade_system_liquidation_cost = "facade_system_liquidation_cost";
        string facade_system_liquidation_cost_lbl = "Cost of liquidation";
        string facade_system_remnant_value = "facade_system_remnant_value";
        string facade_system_remnant_value_lbl = "Remnant value";


        // Windows
        string change_windows = "WINDOWS__CHANGE";
        string change_windows_lbl = "Change windows";
        string windows_life_of_product = "WINDOWS__LIFE_OF_PRODUCT";
        string windows_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string windows_type_windows = "WINDOWS__TYPE_OF_WINDOWS";
        //string windows_type_of_windows_lbl = "Material in frame";
        //string windows_change_in_annual_heat_demand_due_to_windows = "WINDOWS__CHANGE_IN_AHD_DUE_TO_WINDOWS";
        //string windows_change_in_annual_heat_demand_due_to_windows_lbl = "Change in annual heat demand due to windows (an energy saving is given as a negative value)";
        string windows_initial_investment = "windows_initial_investment";
        string windows_initial_investment_lbl = "Initial investment";
        string windows_installation_cost = "windows_installation_cost";
        string windows_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string windows_maintenance_cost = "windows_maintenance_cost";
        string windows_maintenance_cost_lbl = "Total maintenance costs per year";
        string windows_taxes_fees_cost = "windows_taxes_fees_cost";
        string windows_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string windows_liquidation_cost = "windows_liquidation_cost";
        string windows_liquidation_cost_lbl = "Cost of liquidation";
        string windows_remnant_value = "windows_remnant_value";
        string windows_remnant_value_lbl = "Remnant value";

        // Doors
        string change_doors = "DOORS__CHANGE";
        string change_doors_lbl = "Change doors";
        string doors_life_of_product = "DOORS__LIFE_OF_PRODUCT";
        string doors_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string doors_type_doors = "DOORS__TYPE_OF_DOORS";
        //string doors_type_of_doors_lbl = "Type of doors";
        //string doors_change_in_annual_heat_demand_due_to_doors = "DOORS__CHANGE_IN_AHD_DUE_TO_DOORS";
        //string doors_change_in_annual_heat_demand_due_to_doors_lbl = "Change in annual heat demand due to doors (an energy saving is given as a negative value)";
        //string doors_number_of_new_front_doors = "DOORS__NUMBER_OF_NEW_FRONT_DOORS";
        //string doors_number_of_new_front_doors_lbl = "Number of new front doors (required if renovation includes new doors)";
        string doors_initial_investment = "doors_initial_investment";
        string doors_initial_investment_lbl = "Initial investment";
        string doors_installation_cost = "doors_installation_cost";
        string doors_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string doors_maintenance_cost = "doors_maintenance_cost";
        string doors_maintenance_cost_lbl = "Total maintenance costs per year";
        string doors_taxes_fees_cost = "doors_taxes_fees_cost";
        string doors_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string doors_liquidation_cost = "doors_liquidation_cost";
        string doors_liquidation_cost_lbl = "Cost of liquidation";
        string doors_remnant_value = "doors_remnant_value";
        string doors_remnant_value_lbl = "Remnant value";
        #endregion

        #region Ventilation
        // Ventilation
        // Ventilation ducts
        string change_ventilation_ducts = "VENTILATION_DUCTS__CHANGE";
        string change_ventilation_ducts_lbl = "Change ventilation ducts";
        string ventilation_ducts_life_of_product = "VENTILATION_DUCTS__LIFE_OF_PRODUCT";
        string ventilation_ducts_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string ventilation_ducts_type_of_material = "VENTILATION_DUCTS__MATERIAL_OF_VENTILATION_DUCTS";
        //string ventilation_ducts_type_of_material_lbl = "Material in ventilation ducts";
        //string ventilation_ducts_weight_of_ventilation_ducts = "VENTILATION_DUCTS__WEIGHT_OF_VENTILATION_DUCTS";
        //string ventilation_ducts_weight_of_ventilation_ducts_lbl = "Weight of ventilation ducts (Required if renovation includes new ventilation ducts)";
        string ventilation_ducts_initial_investment = "ventilation_ducts_initial_investment";
        string ventilation_ducts_initial_investment_lbl = "Initial investment";
        string ventilation_ducts_installation_cost = "ventilation_ducts_installation_cost";
        string ventilation_ducts_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string ventilation_ducts_maintenance_cost = "ventilation_ducts_maintenance_cost";
        string ventilation_ducts_maintenance_cost_lbl = "Total maintenance costs per year";
        string ventilation_ducts_taxes_fees_cost = "ventilation_ducts_taxes_fees_cost";
        string ventilation_ducts_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string ventilation_ducts_liquidation_cost = "ventilation_ducts_liquidation_cost";
        string ventilation_ducts_liquidation_cost_lbl = "Cost of liquidation";
        string ventilation_ducts_remnant_value = "ventilation_ducts_remnant_value";
        string ventilation_ducts_remnant_value_lbl = "Remnant value";

        // Airflow assembly
        string change_airflow_assembly = "AIR_FLOW_ASSEMBLY__CHANGE";
        string change_airflow_assembly_lbl = "Change airflow assembly";
        string airflow_assembly_life_of_product = "AIR_FLOW_ASSEMBLY__LIFE_OF_PRODUCT";
        string airflow_assembly_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string airflow_assembly_type_of_airflow_assembly = "AIR_FLOW_ASSEMBLY__TYPE_OF_AIR_FLOW_ASSEMBLY";
        //string airflow_assembly_type_of_airflow_assembly_lbl = "Type of airflow assembly";
        string airflow_assembly_initial_investment = "airflow_assembly_initial_investment";
        string airflow_assembly_initial_investment_lbl = "Initial investment";
        string airflow_assembly_installation_cost = "airflow_assembly_installation_cost";
        string airflow_assembly_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string airflow_assembly_heat_consumption = "airflow_assembly_heat_consumption";
        string airflow_assembly_heat_consumption_lbl = "Heat consumption";
        string airflow_assembly_heat_annual_use = "airflow_assembly_heat_annual_use";
        string airflow_assembly_heat_annual_use_lbl = "Annual use";  //Same for all components? (e.g. electric,water and natural gas)
        string airflow_assembly_electric_consumption = "airflow_assembly_electric_consumption";
        string airflow_assembly_electric_consumption_lbl = "Electric consumption";
        string airflow_assembly_electric_annual_use = "airflow_assembly_electric_annual_use";
        string airflow_assembly_electric_annual_use_lbl = "Electric annual use";
        string airflow_assembly_maintenance_cost = "airflow_assembly_maintenance_cost";
        string airflow_assembly_maintenance_cost_lbl = "Total maintenance costs per year";
        string airflow_assembly_taxes_fees_cost = "airflow_assembly_taxes_fees_cost";
        string airflow_assembly_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string airflow_assembly_liquidation_cost = "airflow_assembly_liquidation_cost";
        string airflow_assembly_liquidation_cost_lbl = "Cost of liquidation";
        string airflow_assembly_remnant_value = "airflow_assembly_remnant_value";
        string airflow_assembly_remnant_value_lbl = "Remnant value";

        // Air distribution housings and silencer
        string change_air_distribution_housings_and_silencers = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__CHANGE";
        string change_air_distribution_housings_and_silencers_lbl = "Change air distribution housings and silencers";
        //string air_distribution_housings_and_silencers_number_of_distribution_housings = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__NUMBER_OF_NEW_HOUSINGS";
        //string air_distribution_housings_and_silencers_number_of_distribution_housings_lbl = "Number of air distribution housings";
        string air_distribution_housings_and_silencers_life_of_product = "AIR_DISTRIBUTION_HOUSINGS_AND_SILENCERS__LIFE_OF_PRODUCT";
        string air_distribution_housings_and_silencers_life_of_product_lbl = "Life of air distribution housings and silencers (practical time of life of the products and materials used)";
        string air_distribution_housings_and_silencers_initial_investment = "air_distribution_housings_and_silencers_initial_investment";
        string air_distribution_housings_and_silencers_initial_investment_lbl = "Initial investment";
        string air_distribution_housings_and_silencers_installation_cost = "air_distribution_housings_and_silencers_installation_cost";
        string air_distribution_housings_and_silencers_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string air_distribution_housings_and_silencers_maintenance_cost = "air_distribution_housings_and_silencers_maintenance_cost";
        string air_distribution_housings_and_silencers_maintenance_cost_lbl = "Total maintenance costs per year";
        string air_distribution_housings_and_silencers_taxes_fees_cost = "air_distribution_housings_and_silencers_taxes_fees_cost";
        string air_distribution_housings_and_silencers_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string air_distribution_housings_and_silencers_liquidation_cost = "air_distribution_housings_and_silencers_liquidation_cost";
        string air_distribution_housings_and_silencers_liquidation_cost_lbl = "Cost of liquidation";
        string air_distribution_housings_and_silencers_remnant_value = "air_distribution_housings_and_silencers_remnant_value";
        string air_distribution_housings_and_silencers_remnant_value_lbl = "Remnant value";

        #endregion
        
        #region Radiators, pipes and electricity
        // Radiators, pipes and electricity
        // Radiators
        string change_radiators = "RADIATORS__CHANGE";
        string change_radiators_lbl = "Change radiators";
        string radiators_life_of_product = "RADIATORS__LIFE_OF_PRODUCT";
        string radiators_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string radiators_type_of_radiators = "RADIATORS__TYPE_OF_RADIATORS";
        //string radiators_type_of_radiators_lbl = "Type of radiators";
        //string radiators_weight_of_radiators = "RADIATORS__WEIGHT_OF_NEW_RADIATORS";
        //string radiators_weight_of_radiators_lbl = "Weight of new radiators";
        string radiators_initial_investment = "radiators_initial_investment";
        string radiators_initial_investment_lbl = "Initial investment";
        string radiators_installation_cost = "radiators_installation_cost";
        string radiators_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string radiators_maintenance_cost = "radiators_maintenance_cost";
        string radiators_maintenance_cost_lbl = "Total maintenance costs per year";
        string radiators_taxes_fees_cost = "radiators_taxes_fees_cost";
        string radiators_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string radiators_liquidation_cost = "radiators_liquidation_cost";
        string radiators_liquidation_cost_lbl = "Cost of liquidation";
        string radiators_remnant_value = "radiators_remnant_value";
        string radiators_remnant_value_lbl = "Remnant value";

        // Piping System - Copper
        string change_piping_copper = "PIPING_SYSTEM_COPPER__CHANGE";
        string change_piping_copper_lbl = "Change copper pipes";
        string piping_copper_life_of_product = "PIPING_SYSTEM_COPPER__LIFE_OF_PRODUCT";
        string piping_copper_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string piping_copper_weight_of_copper_pipes = "PIPING_SYSTEM_COPPER__WEIGHT_OF_NEW_PIPES";
        //string piping_copper_weight_of_copper_pipes_lbl = "Weight of new pipes";
        string piping_copper_initial_investment = "piping_copper_initial_investment";
        string piping_copper_initial_investment_lbl = "Initial investment";
        string piping_copper_installation_cost = "piping_copper_installation_cost";
        string piping_copper_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string piping_copper_maintenance_cost = "piping_copper_maintenance_cost";
        string piping_copper_maintenance_cost_lbl = "Total maintenance costs per year";
        string piping_copper_taxes_fees_cost = "piping_copper_taxes_fees_cost";
        string piping_copper_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string piping_copper_liquidation_cost = "piping_copper_liquidation_cost";
        string piping_copper_liquidation_cost_lbl = "Cost of liquidation";
        string piping_copper_remnant_value = "piping_copper_remnant_value";
        string piping_copper_remnant_value_lbl = "Remnant value";

        // Piping System - PEX
        string change_piping_pex = "PIPING_SYSTEM_PEX__CHANGE";
        string change_piping_pex_lbl = "Change PEX pipes";
        string piping_pex_life_of_product = "PIPING_SYSTEM_PEX__LIFE_OF_PRODUCT";
        string piping_pex_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string piping_pex_weight_of_pex_pipes = "PIPING_SYSTEM_PEX__WEIGHT_OF_NEW_PIPES";
        //string piping_pex_weight_of_pex_pipes_lbl = "Weight of new pipes";
        string piping_pex_initial_investment = "piping_pex_initial_investment";
        string piping_pex_initial_investment_lbl = "Initial investment";
        string piping_pex_installation_cost = "piping_pex_installation_cost";
        string piping_pex_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string piping_pex_maintenance_cost = "piping_pex_maintenance_cost";
        string piping_pex_maintenance_cost_lbl = "Total maintenance costs per year";
        string piping_pex_taxes_fees_cost = "piping_pex_taxes_fees_cost";
        string piping_pex_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string piping_pex_liquidation_cost = "piping_pex_liquidation_cost";
        string piping_pex_liquidation_cost_lbl = "Cost of liquidation";
        string piping_pex_remnant_value = "piping_pex_remnant_value";
        string piping_pex_remnant_value_lbl = "Remnant value";

        // Piping System - PP
        string change_piping_pp = "PIPING_SYSTEM_PP__CHANGE";
        string change_piping_pp_lbl = "Change PP pipes";
        string piping_pp_life_of_product = "PIPING_SYSTEM_PP__LIFE_OF_PRODUCT";
        string piping_pp_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string piping_pp_weight_of_pp_pipes = "PIPING_SYSTEM_PP__WEIGHT_OF_NEW_PIPES";
        //string piping_pp_weight_of_pp_pipes_lbl = "Weight of new pipes";
        string piping_pp_initial_investment = "piping_pp_initial_investment";
        string piping_pp_initial_investment_lbl = "Initial investment";
        string piping_pp_installation_cost = "piping_pp_installation_cost";
        string piping_pp_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string piping_pp_maintenance_cost = "piping_pp_maintenance_cost";
        string piping_pp_maintenance_cost_lbl = "Total maintenance costs per year";
        string piping_pp_taxes_fees_cost = "piping_pp_taxes_fees_cost";
        string piping_pp_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string piping_pp_liquidation_cost = "piping_pp_liquidation_cost";
        string piping_pp_liquidation_cost_lbl = "Cost of liquidation";
        string piping_pp_remnant_value = "piping_pp_remnant_value";
        string piping_pp_remnant_value_lbl = "Remnant value";

        // Piping System - Cast Iron
        string change_piping_cast_iron = "PIPING_SYSTEM_CAST_IRON__CHANGE";
        string change_piping_cast_iron_lbl = "Change cast iron pipes";
        string piping_cast_iron_life_of_product = "PIPING_SYSTEM_CAST_IRON__LIFE_OF_PRODUCT";
        string piping_cast_iron_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string piping_cast_iron_weight_of_cast_iron_pipes = "PIPING_SYSTEM_CAST_IRON__WEIGHT_OF_NEW_PIPES";
        //string piping_cast_iron_weight_of_cast_iron_pipes_lbl = "Weight of new pipes";
        string piping_cast_iron_initial_investment = "piping_cast_iron_initial_investment";
        string piping_cast_iron_initial_investment_lbl = "Initial investment";
        string piping_cast_iron_installation_cost = "piping_cast_iron_installation_cost";
        string piping_cast_iron_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string piping_cast_iron_maintenance_cost = "piping_cast_iron_maintenance_cost";
        string piping_cast_iron_maintenance_cost_lbl = "Total maintenance costs per year";
        string piping_cast_iron_taxes_fees_cost = "piping_cast_iron_taxes_fees_cost";
        string piping_cast_iron_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string piping_cast_iron_liquidation_cost = "piping_cast_iron_liquidation_cost";
        string piping_cast_iron_liquidation_cost_lbl = "Cost of liquidation";
        string piping_cast_iron_remnant_value = "piping_cast_iron_remnant_value";
        string piping_cast_iron_remnant_value_lbl = "Remnant value";

        // Piping System - Galvanized Steel
        string change_piping_galvanized_steel = "PIPING_SYSTEM_GALVANISED_STEEL__CHANGE";
        string change_piping_galvanized_steel_lbl = "Change galvanized steel pipes";
        string piping_galvanized_steel_life_of_product = "PIPING_SYSTEM_GALVANISED_STEEL__LIFE_OF_PRODUCT";
        string piping_galvanized_steel_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string piping_galvanized_steel_weight_of_galvanized_steel_pipes = "PIPING_SYSTEM_GALVANISED_STEEL__WEIGHT_OF_NEW_PIPES";
        //string piping_galvanized_steel_weight_of_galvanized_steel_pipes_lbl = "Weight of new pipes";
        string piping_galvanized_steel_initial_investment = "piping_galvanized_steel_initial_investment";
        string piping_galvanized_steel_initial_investment_lbl = "Initial investment";
        string piping_galvanized_steel_installation_cost = "piping_galvanized_steel_installation_cost";
        string piping_galvanized_steel_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string piping_galvanized_steel_maintenance_cost = "piping_galvanized_steel_maintenance_cost";
        string piping_galvanized_steel_maintenance_cost_lbl = "Total maintenance costs per year";
        string piping_galvanized_steel_taxes_fees_cost = "piping_galvanized_steel_taxes_fees_cost";
        string piping_galvanized_steel_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string piping_galvanized_steel_liquidation_cost = "piping_galvanized_steel_liquidation_cost";
        string piping_galvanized_steel_liquidation_cost_lbl = "Cost of liquidation";
        string piping_galvanized_steel_remnant_value = "piping_galvanized_steel_remnant_value";
        string piping_galvanized_steel_remnant_value_lbl = "Remnant value";

        // Piping System - Relining
        string change_piping_relining = "PIPING_SYSTEM_RELINING__CHANGE";
        string change_piping_relining_lbl = "Relining of pipes";
        string piping_relining_life_of_product = "PIPING_SYSTEM_RELINING__LIFE_OF_PRODUCT";
        string piping_relining_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string piping_relining_weight_of_relining_pipes = "PIPING_SYSTEM_RELINING__WEIGHT_OF_NEW_PIPES";
        //string piping_relining_weight_of_relining_pipes_lbl = "Weight of new pipes";
        string piping_relining_initial_investment = "piping_relining_initial_investment";
        string piping_relining_initial_investment_lbl = "Initial investment";
        string piping_relining_installation_cost = "piping_relining_installation_cost";
        string piping_relining_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string piping_relining_maintenance_cost = "piping_relining_maintenance_cost";
        string piping_relining_maintenance_cost_lbl = "Total maintenance costs per year";
        string piping_relining_taxes_fees_cost = "piping_relining_taxes_fees_cost";
        string piping_relining_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string piping_relining_liquidation_cost = "piping_relining_liquidation_cost";
        string piping_relining_liquidation_cost_lbl = "Cost of liquidation";
        string piping_relining_remnant_value = "piping_relining_remnant_value";
        string piping_relining_remnant_value_lbl = "Remnant value";

        // Electrical wiring
        string change_electrical_wiring = "ELECTRICAL_WIRING__CHANGE";
        string change_electrical_wiring_lbl = "Replace electrical wiring";
        string electrical_wiring_life_of_product = "ELECTRICAL_WIRING__LIFE_OF_PRODUCT";
        string electrical_wiring_life_of_product_lbl = "Life of product (practical time of life of the products and materials used)";
        //string electrical_wiring_weight_of_electrical_wiring = "ELECTRICAL_WIRING__WEIGHT_OF_NEW_WIRES";
        //string electrical_wiring_weight_of_electrical_wiring_lbl = "Weight of new wires";
        string electrical_wiring_initial_investment = "electrical_wiring_initial_investment";
        string electrical_wiring_initial_investment_lbl = "Initial investment";
        string electrical_wiring_installation_cost = "electrical_wiring_installation_cost";
        string electrical_wiring_installation_cost_lbl = "Installation cost (including possible cost of delivery)";
        string electrical_wiring_maintenance_cost = "electrical_wiring_maintenance_cost";
        string electrical_wiring_maintenance_cost_lbl = "Total maintenance costs per year";
        string electrical_wiring_taxes_fees_cost = "electrical_wiring_taxes_fees_cost";
        string electrical_wiring_taxes_fees_cost_lbl = "Taxes / Fees per year";
        string electrical_wiring_liquidation_cost = "electrical_wiring_liquidation_cost";
        string electrical_wiring_liquidation_cost_lbl = "Cost of liquidation";
        string electrical_wiring_remnant_value = "electrical_wiring_remnant_value";
        string electrical_wiring_remnant_value_lbl = "Remnant value";

        #endregion

        #endregion

        #endregion

        #endregion

        public LCC_Module()
        {
            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "MSR_LCC";

            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_lcc };

            //Error handler
            this.ErrorRaised += LCC_Module_ErrorRaised;

            //Notification
            this.StatusMessage += LCC_Module_StatusMessage;

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

        void LCC_Module_StatusMessage(object sender, StatusEventArg e)
        {
            Console.WriteLine(String.Format("Status message:\n\t{0}", e.StatusMessage));
        }

        void LCC_Module_ErrorRaised(object sender, ErrorMessageEventArg e)
        {
            Console.WriteLine(String.Format("Error message: {0}", e.Message));
            if (e.SourceFunction != null & e.SourceFunction != "")
                Console.WriteLine(String.Format("\tIn source function: {0}", e.SourceFunction));
        }

        void LCC_Module_ErrorRaised(object sender, Exception ex)
        {
            ErrorMessageEventArg em = new ErrorMessageEventArg();
            em.Message = ex.Message;
            LCC_Module_ErrorRaised(sender, em);
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
            ///commonProp.Add(lcc_calculation_period, new Number(label: "Period of time for which total life cycle impact is summarized", min: 1, unit: "years", order: ++order));
            commonProp.Add(discount_rate, new Number(label: discount_rate_lbl, value: 6, unit: "%", order: ++order));
            commonProp.Add(electric_cost, new Number(label: electric_cost_lbl, value: 0.208, unit: "EUR/kWh", order: ++order));
            commonProp.Add(heat_cost, new Number(label: heat_cost_lbl, value: 0.06, unit: "EUR/kWh", order: ++order));
            commonProp.Add(water_cost, new Number(label: water_cost_lbl, value: 1.91, unit: "EUR/1000 liters", order: ++order));
            commonProp.Add(natural_gas_cost, new Number(label: natural_gas_cost_lbl, value: 0.072, unit: "EUR/kWh", order: ++order));
            // If district heating is used (before/after renovation)
            //commonProp.Add(key: ep_district, item: new Number(label: ep_district_lbl, min: 0, unit: "EUR", order: ++order));
            //commonProp.Add(key: peu_district, item: new Number(label: peu_district_lbl, min: 0, unit: "kWh/kWh", order: ++order));

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
            //intstr += "You need to fill in the  building properties as well as the parameters under checked checkboxes. ";
            //intstr += "If this is the as-is step leave all checkboxes unchecked. ";
            intstr += "If multiple buildings have common properties you may select those buildings and assign them values simultaneously. ";
            InputGroup instructions = new InputGroup(label: intstr, order: ++order);
            buildning_specific_data.Add(key: "instructions", item: instructions);

            // Building Common
            //++order;
            //BuildingProperties(ref buildning_specific_data, ref order);

            // Heating System
            ++order;
            HeatingSystem(ref buildning_specific_data, ref order);

            //// Building Shell
            ++order;
            BuildingShell(ref buildning_specific_data, ref order);

            // Ventilation System
            ++order;
            VentilationSystem(ref buildning_specific_data, ref order);

            //// Radiators, pipes and electricity
            ++order;
            RadiatorsPipesElectricity(ref buildning_specific_data, ref order);

            return buildning_specific_data;
        }

        void BuildingProperties(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("building_properties", new InputGroup("Building Properties", order: ++order));

            // Inputs required in all cases
            //input.Add(key: heat_source_before, item: new Select(label: heat_source_before_lbl, options: heat_sources, value: heat_sources.Last(), order: ++order));
            //input.Add(key: heated_area, item: new Number(label: heated_area_lbl, min: 1, unit: "m\u00b2", order: ++order, value: 99));
            //input.Add(key: nr_apartments, item: new Number(label: nr_apartments_lbl, min: 1, order: ++order, value: 98));


        }

        void HeatingSystem(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("heating_system", new InputGroup("Renovate Heating System", ++order));

            // Change Heating System
            input.Add(key: change_heating_system, item: new Checkbox(label: change_heating_system_lbl, value: false, order: ++order));
            input.Add(key: heating_system_life_of_product, item: new Number(label: heating_system_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: heating_system_initial_investment, item: new Number(label: heating_system_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: heating_system_installation_cost, item: new Number(label: heating_system_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: heating_system_heat_consumption, item: new Number(label: heating_system_heat_consumption_lbl, min: 0, unit: "kWh", order: ++order));
            input.Add(key: heating_system_heat_annual_use, item: new Number(label: heating_system_heat_annual_use_lbl, min: 0, unit: "hours", order: ++order));
            input.Add(key: heating_system_water_consumption, item: new Number(label: heating_system_water_consumption_lbl, min: 0, unit: "1000 liters", order: ++order));
            input.Add(key: heating_system_water_annual_use, item: new Number(label: heating_system_water_annual_use_lbl, min: 0, unit: "hours", order: ++order));
            input.Add(key: heating_system_electric_consumption, item: new Number(label: heating_system_electric_consumption_lbl, min: 0, unit: "kWh", order: ++order));
            input.Add(key: heating_system_electric_annual_use, item: new Number(label: heating_system_electric_annual_use_lbl, min: 0, unit: "hours", order: ++order));
            input.Add(key: heating_system_natural_gas_consumption, item: new Number(label: heating_system_natural_gas_consumption_lbl, min: 0, unit: "kWh", order: ++order));
            input.Add(key: heating_system_natural_gas_annual_use, item: new Number(label: heating_system_natural_gas_annual_use_lbl, min: 0, unit: "hours", order: ++order));
            input.Add(key: heating_system_maintenance_cost, item: new Number(label: heating_system_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: heating_system_taxes_fees_cost, item: new Number(label: heating_system_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: heating_system_liquidation_cost, item: new Number(label: heating_system_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: heating_system_remnant_value, item: new Number(label: heating_system_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

            // Change Circulation Pump
            input.Add(key: change_circulationpump_in_heating_system, item: new Checkbox(label: change_circulationpump_in_heating_system_lbl, value: false, order: ++order));
            input.Add(key: circulationpump_life_of_product, item: new Number(label: circulationpump_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: circulationpump_initial_investment, item: new Number(label: circulationpump_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: circulationpump_installation_cost, item: new Number(label: circulationpump_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: circulationpump_electric_consumption, item: new Number(label: circulationpump_electric_consumption_lbl, min: 0, unit: "kWh", order: ++order));
            input.Add(key: circulationpump_electric_annual_use, item: new Number(label: circulationpump_electric_annual_use_lbl, min: 0, unit: "hours", order: ++order));
            input.Add(key: circulationpump_maintenance_cost, item: new Number(label: circulationpump_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: circulationpump_taxes_fees_cost, item: new Number(label: circulationpump_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: circulationpump_liquidation_cost, item: new Number(label: circulationpump_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: circulationpump_remnant_value, item: new Number(label: circulationpump_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));
        }

        void BuildingShell(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("building_shell", new InputGroup("Renovate Building Shell", ++order));

            // Insulation material 1
            input.Add(key: change_insulation_material_1, item: new Checkbox(label: change_insulation_material_1_lbl, order: ++order));
            //input.Add(key: insulation_material_1_type_of_insulation, item: new Select(label: insulation_material_1_type_of_insulation_lbl, options: type_of_insulation, order: ++order));
            input.Add(key: insulation_material_1_life_of_product, item: new Number(label: insulation_material_1_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: insulation_material_1_amount_of_new_insulation_material, item: new Number(label: insulation_material_1_amount_of_new_insulation_material_lbl, min: 0, unit: "kg", order: ++order));
            input.Add(key: insulation_material_1_initial_investment, item: new Number(label: insulation_material_1_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: insulation_material_1_installation_cost, item: new Number(label: insulation_material_1_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: insulation_material_1_maintenance_cost, item: new Number(label: insulation_material_1_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: insulation_material_1_taxes_fees_cost, item: new Number(label: insulation_material_1_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: insulation_material_1_liquidation_cost, item: new Number(label: insulation_material_1_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: insulation_material_1_remnant_value, item: new Number(label: insulation_material_1_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

            // Insulation material 2
            input.Add(key: change_insulation_material_2, item: new Checkbox(label: change_insulation_material_2_lbl, order: ++order));
            //input.Add(key: insulation_material_2_type_of_insulation, item: new Select(label: insulation_material_2_type_of_insulation_lbl, options: type_of_insulation, order: ++order));
            input.Add(key: insulation_material_2_life_of_product, item: new Number(label: insulation_material_2_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: insulation_material_2_amount_of_new_insulation_material, item: new Number(label: insulation_material_2_amount_of_new_insulation_material_lbl, min: 0, unit: "kg", order: ++order));
            input.Add(key: insulation_material_2_initial_investment, item: new Number(label: insulation_material_2_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: insulation_material_2_installation_cost, item: new Number(label: insulation_material_2_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: insulation_material_2_maintenance_cost, item: new Number(label: insulation_material_2_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: insulation_material_2_taxes_fees_cost, item: new Number(label: insulation_material_2_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: insulation_material_2_liquidation_cost, item: new Number(label: insulation_material_2_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: insulation_material_2_remnant_value, item: new Number(label: insulation_material_2_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

            // facade System
            input.Add(key: change_facade_system, item: new Checkbox(label: change_facade_system_lbl, order: ++order));
            //input.Add(key: facade_system_type_facade_system, item: new Select(label: facade_system_type_of_facade_system_lbl, options: type_of_facade_system, order: ++order));
            input.Add(key: facade_system_life_of_product, item: new Number(label: facade_system_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: facade_system_area_of_new_facade_system, item: new Number(label: facade_system_area_of_new_facade_system_lbl, min: 0, unit: "m\u00b2", order: ++order));
            input.Add(key: facade_system_initial_investment, item: new Number(label: facade_system_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: facade_system_installation_cost, item: new Number(label: facade_system_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: facade_system_maintenance_cost, item: new Number(label: facade_system_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: facade_system_taxes_fees_cost, item: new Number(label: facade_system_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: facade_system_liquidation_cost, item: new Number(label: facade_system_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: facade_system_remnant_value, item: new Number(label: facade_system_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

            // Windows
            input.Add(key: change_windows, item: new Checkbox(label: change_windows_lbl, order: ++order));
            //input.Add(key: windows_type_windows, item: new Select(label: windows_type_of_windows_lbl, options: type_of_windows, order: ++order));
            input.Add(key: windows_life_of_product, item: new Number(label: windows_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: windows_area_of_new_windows, item: new Number(label: windows_area_of_new_windows_lbl, min: 0, unit: "m\u00b2", order: ++order));
            input.Add(key: windows_initial_investment, item: new Number(label: windows_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: windows_installation_cost, item: new Number(label: windows_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: windows_maintenance_cost, item: new Number(label: windows_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: windows_taxes_fees_cost, item: new Number(label: windows_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: windows_liquidation_cost, item: new Number(label: windows_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: windows_remnant_value, item: new Number(label: windows_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));
            
            // Doors
            input.Add(key: change_doors, item: new Checkbox(label: change_doors_lbl, order: ++order));
            //input.Add(key: doors_type_doors, item: new Select(label: doors_type_of_doors_lbl, options: type_of_doors, order: ++order));
            input.Add(key: doors_life_of_product, item: new Number(label: doors_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: doors_number_of_new_front_doors, item: new Number(label: doors_number_of_new_front_doors_lbl, min: 0, order: ++order));
            input.Add(key: doors_initial_investment, item: new Number(label: doors_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: doors_installation_cost, item: new Number(label: doors_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: doors_maintenance_cost, item: new Number(label: doors_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: doors_taxes_fees_cost, item: new Number(label: doors_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: doors_liquidation_cost, item: new Number(label: doors_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: doors_remnant_value, item: new Number(label: doors_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

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
            //input.Add(key: ventilation_ducts_type_of_material, item: new Select(label: ventilation_ducts_type_of_material_lbl, options: type_of_ventilation_ducts_material, order: ++order));
            input.Add(key: ventilation_ducts_life_of_product, item: new Number(label: ventilation_ducts_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: ventilation_ducts_weight_of_ventilation_ducts, item: new Number(label: ventilation_ducts_weight_of_ventilation_ducts_lbl, unit: "kWh/year", order: ++order));
            input.Add(key: ventilation_ducts_initial_investment, item: new Number(label: ventilation_ducts_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: ventilation_ducts_installation_cost, item: new Number(label: ventilation_ducts_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: ventilation_ducts_maintenance_cost, item: new Number(label: ventilation_ducts_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: ventilation_ducts_taxes_fees_cost, item: new Number(label: ventilation_ducts_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: ventilation_ducts_liquidation_cost, item: new Number(label: ventilation_ducts_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: ventilation_ducts_remnant_value, item: new Number(label: ventilation_ducts_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

            // Airflow assembly
            input.Add(key: change_airflow_assembly, item: new Checkbox(label: change_airflow_assembly_lbl, order: ++order));
            //input.Add(key: airflow_assembly_type_of_airflow_assembly, item: new Select(label: airflow_assembly_type_of_airflow_assembly_lbl, options: type_of_airflow_assembly, order: ++order));
            input.Add(key: airflow_assembly_life_of_product, item: new Number(label: airflow_assembly_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: airflow_assembly_design_airflow_exhaust_air, item: new Number(label: airflow_assembly_design_airflow_exhaust_air_lbl, unit: "kWh/year", order: ++order));
            input.Add(key: airflow_assembly_initial_investment, item: new Number(label: airflow_assembly_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: airflow_assembly_installation_cost, item: new Number(label: airflow_assembly_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: airflow_assembly_heat_consumption, item: new Number(label: airflow_assembly_heat_consumption_lbl, min: 0, unit: "kWh", order: ++order));
            input.Add(key: airflow_assembly_heat_annual_use, item: new Number(label: airflow_assembly_heat_annual_use_lbl, min: 0, unit: "hours", order: ++order));
            input.Add(key: airflow_assembly_electric_consumption, item: new Number(label: airflow_assembly_electric_consumption_lbl, min: 0, unit: "kWh", order: ++order));
            input.Add(key: airflow_assembly_electric_annual_use, item: new Number(label: airflow_assembly_electric_annual_use_lbl, min: 0, unit: "hours", order: ++order));
            input.Add(key: airflow_assembly_maintenance_cost, item: new Number(label: airflow_assembly_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: airflow_assembly_taxes_fees_cost, item: new Number(label: airflow_assembly_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: airflow_assembly_liquidation_cost, item: new Number(label: airflow_assembly_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: airflow_assembly_remnant_value, item: new Number(label: airflow_assembly_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

            // Air distribution housings and silencers
            input.Add(key: change_air_distribution_housings_and_silencers, item: new Checkbox(label: change_air_distribution_housings_and_silencers_lbl, order: ++order));
            //input.Add(key: air_distribution_housings_and_silencers_number_of_distribution_housings, item: new Number(label: air_distribution_housings_and_silencers_number_of_distribution_housings_lbl, min: 0, order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_life_of_product, item: new Number(label: air_distribution_housings_and_silencers_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_initial_investment, item: new Number(label: air_distribution_housings_and_silencers_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_installation_cost, item: new Number(label: air_distribution_housings_and_silencers_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_maintenance_cost, item: new Number(label: air_distribution_housings_and_silencers_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_taxes_fees_cost, item: new Number(label: air_distribution_housings_and_silencers_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_liquidation_cost, item: new Number(label: air_distribution_housings_and_silencers_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: air_distribution_housings_and_silencers_remnant_value, item: new Number(label: air_distribution_housings_and_silencers_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));

        }

        void RadiatorsPipesElectricity(ref GeoJson input, ref int order)
        {
            //Header
            input.Add("radiators_pipes_and_electricity", new InputGroup("Renovate Radiators, Pipes and/or Electricity", ++order));

            // Radiators
            input.Add(key: change_radiators, item: new Checkbox(label: change_radiators_lbl, order: ++order));
            //input.Add(key: radiators_type_of_radiators, item: new Select(label: radiators_type_of_radiators_lbl, options: type_of_radiators, order: ++order));
            input.Add(key: radiators_life_of_product, item: new Number(label: radiators_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: radiators_weight_of_radiators, item: new Number(label: radiators_weight_of_radiators_lbl, unit: "kg", order: ++order));
            input.Add(key: radiators_initial_investment, item: new Number(label: radiators_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: radiators_installation_cost, item: new Number(label: radiators_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: radiators_maintenance_cost, item: new Number(label: radiators_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: radiators_taxes_fees_cost, item: new Number(label: radiators_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: radiators_liquidation_cost, item: new Number(label: radiators_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: radiators_remnant_value, item: new Number(label: radiators_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


            // Piping System - Copper
            input.Add(key: change_piping_copper, item: new Checkbox(label: change_piping_copper_lbl, order: ++order));
            input.Add(key: piping_copper_life_of_product, item: new Number(label: piping_copper_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: piping_copper_weight_of_copper_pipes, item: new Number(label: piping_copper_weight_of_copper_pipes_lbl, unit: "kg", order: ++order));
            input.Add(key: piping_copper_initial_investment, item: new Number(label: piping_copper_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_copper_installation_cost, item: new Number(label: piping_copper_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_copper_maintenance_cost, item: new Number(label: piping_copper_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_copper_taxes_fees_cost, item: new Number(label: piping_copper_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_copper_liquidation_cost, item: new Number(label: piping_copper_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_copper_remnant_value, item: new Number(label: piping_copper_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


            // Piping System - PEX
            input.Add(key: change_piping_pex, item: new Checkbox(label: change_piping_pex_lbl, order: ++order));
            input.Add(key: piping_pex_life_of_product, item: new Number(label: piping_pex_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: piping_pex_weight_of_pex_pipes, item: new Number(label: piping_pex_weight_of_pex_pipes_lbl, unit: "kg", order: ++order));
            input.Add(key: piping_pex_initial_investment, item: new Number(label: piping_pex_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_pex_installation_cost, item: new Number(label: piping_pex_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_pex_maintenance_cost, item: new Number(label: piping_pex_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_pex_taxes_fees_cost, item: new Number(label: piping_pex_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_pex_liquidation_cost, item: new Number(label: piping_pex_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_pex_remnant_value, item: new Number(label: piping_pex_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


            // Piping System - PP
            input.Add(key: change_piping_pp, item: new Checkbox(label: change_piping_pp_lbl, order: ++order));
            input.Add(key: piping_pp_life_of_product, item: new Number(label: piping_pp_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: piping_pp_weight_of_pp_pipes, item: new Number(label: piping_pp_weight_of_pp_pipes_lbl, unit: "kg", order: ++order));
            input.Add(key: piping_pp_initial_investment, item: new Number(label: piping_pp_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_pp_installation_cost, item: new Number(label: piping_pp_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_pp_maintenance_cost, item: new Number(label: piping_pp_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_pp_taxes_fees_cost, item: new Number(label: piping_pp_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_pp_liquidation_cost, item: new Number(label: piping_pp_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_pp_remnant_value, item: new Number(label: piping_pp_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


            // Piping System - Cast Iron
            input.Add(key: change_piping_cast_iron, item: new Checkbox(label: change_piping_cast_iron_lbl, order: ++order));
            input.Add(key: piping_cast_iron_life_of_product, item: new Number(label: piping_cast_iron_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: piping_cast_iron_weight_of_cast_iron_pipes, item: new Number(label: piping_cast_iron_weight_of_cast_iron_pipes_lbl, unit: "kg", order: ++order));
            input.Add(key: piping_cast_iron_initial_investment, item: new Number(label: piping_cast_iron_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_cast_iron_installation_cost, item: new Number(label: piping_cast_iron_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_cast_iron_maintenance_cost, item: new Number(label: piping_cast_iron_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_cast_iron_taxes_fees_cost, item: new Number(label: piping_cast_iron_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_cast_iron_liquidation_cost, item: new Number(label: piping_cast_iron_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_cast_iron_remnant_value, item: new Number(label: piping_cast_iron_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


            // Piping System - Galvanized Steel
            input.Add(key: change_piping_galvanized_steel, item: new Checkbox(label: change_piping_galvanized_steel_lbl, order: ++order));
            input.Add(key: piping_galvanized_steel_life_of_product, item: new Number(label: piping_galvanized_steel_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: piping_galvanized_steel_weight_of_galvanized_steel_pipes, item: new Number(label: piping_galvanized_steel_weight_of_galvanized_steel_pipes_lbl, unit: "kg", order: ++order));
            input.Add(key: piping_galvanized_steel_initial_investment, item: new Number(label: piping_galvanized_steel_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_galvanized_steel_installation_cost, item: new Number(label: piping_galvanized_steel_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_galvanized_steel_maintenance_cost, item: new Number(label: piping_galvanized_steel_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_galvanized_steel_taxes_fees_cost, item: new Number(label: piping_galvanized_steel_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_galvanized_steel_liquidation_cost, item: new Number(label: piping_galvanized_steel_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_galvanized_steel_remnant_value, item: new Number(label: piping_galvanized_steel_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


            // Piping System - Relining
            input.Add(key: change_piping_relining, item: new Checkbox(label: change_piping_relining_lbl, order: ++order));
            input.Add(key: piping_relining_life_of_product, item: new Number(label: piping_relining_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: piping_relining_weight_of_relining_pipes, item: new Number(label: piping_relining_weight_of_relining_pipes_lbl, unit: "kg", order: ++order));
            input.Add(key: piping_relining_initial_investment, item: new Number(label: piping_relining_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_relining_installation_cost, item: new Number(label: piping_relining_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_relining_maintenance_cost, item: new Number(label: piping_relining_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_relining_taxes_fees_cost, item: new Number(label: piping_relining_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: piping_relining_liquidation_cost, item: new Number(label: piping_relining_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: piping_relining_remnant_value, item: new Number(label: piping_relining_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


            // Electrical wiring
            input.Add(key: change_electrical_wiring, item: new Checkbox(label: change_electrical_wiring_lbl, order: ++order));
            input.Add(key: electrical_wiring_life_of_product, item: new Number(label: electrical_wiring_life_of_product_lbl, min: 0, unit: "years", order: ++order));
            //input.Add(key: electrical_wiring_weight_of_electrical_wiring, item: new Number(label: electrical_wiring_weight_of_electrical_wiring_lbl, unit: "kg", order: ++order));
            input.Add(key: electrical_wiring_initial_investment, item: new Number(label: electrical_wiring_initial_investment_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: electrical_wiring_installation_cost, item: new Number(label: electrical_wiring_installation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: electrical_wiring_maintenance_cost, item: new Number(label: electrical_wiring_maintenance_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: electrical_wiring_taxes_fees_cost, item: new Number(label: electrical_wiring_taxes_fees_cost_lbl, min: 0, unit: "EUR/year", order: ++order));
            input.Add(key: electrical_wiring_liquidation_cost, item: new Number(label: electrical_wiring_liquidation_cost_lbl, min: 0, unit: "EUR", order: ++order));
            input.Add(key: electrical_wiring_remnant_value, item: new Number(label: electrical_wiring_remnant_value_lbl, min: 0, unit: "EUR", order: ++order));


        }

        double SetInputDataOneBuilding(Feature building, ref CExcel exls)
        {

            double res = 0.0;

            //SetBuildingProperties(building, ref exls);
            res += SetHeatingSystem(building, ref exls);
            res += SetBuildingShell(building, ref exls);
            res += SetVentilationSystem(building, ref exls);
            res += SetRadiatorsPipesElectricity(building, ref exls);

            return res;

        }

        void SetBuildingProperties(Feature building, ref CExcel exls)
        {
            String Key;
            object value;
            String cell;

            // Inputs required in all cases
            //#region Heated Area
            //Key = heated_area;
            //value = Convert.ToDouble(building.properties[Key]);
            //cell = "C25";
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Number of Apartments
            //Key = nr_apartments;
            //cell = "C26";
            //value = Convert.ToDouble(building.properties[Key]);
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

            //#region Heat Source Before
            //Key = heat_source_before;
            //cell = "C93";
            //value = heat_sources.GetIndex((string)building.properties[Key]) + 1;
            //if (!exls.SetCellValue("Indata", cell, value))
            //    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
            //#endregion

        }

        double SetHeatingSystem(Feature building, ref CExcel exls)
        {
            String Key;
            object value;
            String cell;
            double res = 0.0;

            // Change Heating System
            #region Change Heating System
            Key = change_heating_system;
            value = (bool)building.properties[Key];
            if ((bool)value)
            {
                #region Heating System: Life of Product
                Key = heating_system_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Heating System: Initial Investment
                Key = heating_system_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Heating System: Total Installation Cost
                Key = heating_system_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Heating System: Operating Cost
                #region Heat
                Key = heating_system_heat_consumption;
                cell = "E6";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                Key = heating_system_heat_annual_use;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Water
                Key = heating_system_water_consumption;
                cell = "E11";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                Key = heating_system_water_annual_use;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Electric
                Key = heating_system_electric_consumption;
                cell = "E16";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                Key = heating_system_electric_annual_use;
                cell = "E17";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Natural gas
                Key = heating_system_natural_gas_consumption;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                Key = heating_system_natural_gas_annual_use;
                cell = "E22";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion
                #endregion

                #region Heating System: Maintenace cost
                Key = heating_system_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Heating System: Taxes amd fees cost
                Key = heating_system_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Heating System: Liquidation cost
                Key = heating_system_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Heating System: Remnant value
                Key = heating_system_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }

            #endregion

            // Change Circulation Pump
            #region Change Circulation Pump
            Key = change_circulationpump_in_heating_system;
            cell = "C113";
            value = (bool)(building.properties[Key]);
            if ((bool)value)
            {
                #region Circulation Pump: Life of Product
                Key = circulationpump_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Circulation Pump: Initial Investment
                Key = circulationpump_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Circulation Pump: Total Installation Cost
                Key = circulationpump_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Circulation Pump: Operating Cost
                #region Heat
                cell = "E6";
                if (!exls.SetCellValue("Operating costs", cell, 0))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, 0));

                Key = heating_system_heat_annual_use;
                cell = "E7";
                if (!exls.SetCellValue("Operating costs", cell, 0))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, 0));
                #endregion

                #region Water
                cell = "E11";
                if (!exls.SetCellValue("Operating costs", cell, 0))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, 0));

                cell = "E12";
                if (!exls.SetCellValue("Operating costs", cell, 0))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, 0));
                #endregion

                #region Electric
                Key = circulationpump_electric_consumption;
                cell = "E16";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                Key = circulationpump_electric_annual_use;
                cell = "E17";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                if (!exls.SetCellValue("Operating costs", cell, 0))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, 0));

                cell = "E22";
                if (!exls.SetCellValue("Operating costs", cell, 0))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, 0));
                #endregion
                #endregion

                #region Circulation Pump: Maintenace cost
                Key = circulationpump_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Circulation Pump: Taxes and fess cost
                Key = circulationpump_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Circulation Pump: Liquidation cost
                Key = circulationpump_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Circulation Pump: Ramnant value
                Key = circulationpump_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }

            #endregion

            return res;

        }

        double SetBuildingShell(Feature building, ref CExcel exls)
        {

            String Key;
            object value;
            String cell;
            double res = 0.0;

            // Insulation Material 1
            #region Change Insulation Material 1
            Key = change_insulation_material_1;
            value = (bool)building.properties[Key];
            if ((bool)value)
            {
                #region Insulation Material 1: Life of Product
                Key = insulation_material_1_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Insulation Material 1: Initial Investment
                Key = insulation_material_1_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Insulation Material 1: Total Installation Cost
                Key = insulation_material_1_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Insulation Material 1: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion
                #endregion

                #region Insulation Material 1: Maintenace cost
                Key = insulation_material_1_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Insulation Material 1: Taxes amd fees cost
                Key = insulation_material_1_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Insulation Material 1: Liquidation cost
                Key = insulation_material_1_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Insulation Material 1: Remnant value
                Key = insulation_material_1_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion                      

            // Insulation material 2
            #region Change Insulation Material 2
            Key = change_insulation_material_2;
            value = (bool)building.properties[Key];
            if ((bool)value)
            {
                #region Insulation Material 2: Life of Product
                Key = insulation_material_2_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Insulation Material 2: Initial Investment
                Key = insulation_material_2_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Insulation Material 2: Total Installation Cost
                Key = insulation_material_2_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Insulation Material 2: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Insulation Material 2: Maintenace cost
                Key = insulation_material_2_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Insulation Material 2: Taxes amd fees cost
                Key = insulation_material_2_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Insulation Material 2: Liquidation cost
                Key = insulation_material_2_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Insulation Material 2: Remnant value
                Key = insulation_material_2_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion                      

            // Facade System   
            #region Change Facade System
            Key = change_facade_system;
            value = (bool)building.properties[Key];
            if ((bool)value)
            {
                #region Facade System: Life of Product
                Key = facade_system_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Facade System: Initial Investment
                Key = facade_system_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Facade System: Total Installation Cost
                Key = facade_system_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Facade System: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Facade System: Maintenace cost
                Key = facade_system_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Facade System: Taxes amd fees cost
                Key = facade_system_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Facade System: Liquidation cost
                Key = facade_system_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Facade System: Remnant value
                Key = facade_system_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion 

            // Windows
            #region Change Windows
            Key = change_windows;
            value = (bool)building.properties[Key];
            if ((bool)value)
            {
                #region Windows: Life of Product
                Key = windows_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Windows: Initial Investment
                Key = windows_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Windows: Total Installation Cost
                Key = windows_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Windows: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Windows: Maintenace cost
                Key = windows_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Windows: Taxes amd fees cost
                Key = windows_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Windows: Liquidation cost
                Key = windows_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Windows: Remnant value
                Key = windows_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion 

            // Doors
            #region Change Doors
            Key = change_doors;
            value = (bool)building.properties[Key];
            if ((bool)value)
            {
                #region Doors: Life of Product
                Key = doors_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Doors: Initial Investment
                Key = doors_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Doors: Total Installation Cost
                Key = doors_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Doors: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Doors: Maintenace cost
                Key = doors_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Doors: Taxes amd fees cost
                Key = doors_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Doors: Liquidation cost
                Key = doors_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Doors: Remnant value
                Key = doors_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion 

            return res;

        }

        double SetVentilationSystem(Feature building, ref CExcel exls)
        {
            String Key;
            object value;
            String cell;
            double res = 0;

            // Ventilation Ducts
            #region Change Ventilation Ducts
            #region Ventilation Ducts?
            Key = change_ventilation_ducts;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Ventilation Ducts: Life of Product
                Key = ventilation_ducts_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Ventilation Ducts: Initial Investment
                Key = ventilation_ducts_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Ventilation Ducts: Total Installation Cost
                Key = ventilation_ducts_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Ventilation Ducts: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Ventilation Ducts: Maintenace cost
                Key = ventilation_ducts_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Ventilation Ducts: Taxes amd fees cost
                Key = ventilation_ducts_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Ventilation Ducts: Liquidation cost
                Key = ventilation_ducts_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Ventilation Ducts: Remnant value
                Key = ventilation_ducts_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Airflow Assembly
            #region Change Airflow Assembly
            #region Change Airflow Assembly?
            Key = change_airflow_assembly;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Airflow Assembly: Life of Product
                Key = airflow_assembly_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Airflow Assembly: Initial Investment
                Key = airflow_assembly_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Airflow Assembly: Total Installation Cost
                Key = airflow_assembly_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Airflow Assembly: Operating Cost
                #region Heat
                Key = airflow_assembly_heat_consumption;
                cell = "E6";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                Key = airflow_assembly_heat_annual_use;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Electric
                Key = airflow_assembly_electric_consumption;
                cell = "E16";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                Key = airflow_assembly_electric_annual_use;
                cell = "E17";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion
                #endregion

                #region Airflow Assembly: Maintenace cost
                Key = airflow_assembly_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Airflow Assembly: Taxes amd fees cost
                Key = airflow_assembly_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Airflow Assembly: Liquidation cost
                Key = airflow_assembly_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                #region Airflow Assembly: Remnant value
                Key = airflow_assembly_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {1}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Air Distribution Housings & Silencers
            #region Change Air Distribution Housings & Silencers
            #region Change Air Distribution Housings & Silencers?
            Key = change_air_distribution_housings_and_silencers;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Air Distribution Housings and Silencers: Life of Product
                Key = air_distribution_housings_and_silencers_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Air Distribution Housings and Silencers: Initial Investment
                Key = air_distribution_housings_and_silencers_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Air Distribution Housings and Silencers: Total Installation Cost
                Key = air_distribution_housings_and_silencers_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Air Distribution Housings and Silencers: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Air Distribution Housings and Silencers: Maintenace cost
                Key = air_distribution_housings_and_silencers_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Air Distribution Housings and Silencers: Taxes amd fees cost
                Key = air_distribution_housings_and_silencers_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Air Distribution Housings and Silencers: Liquidation cost
                Key = air_distribution_housings_and_silencers_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Air Distribution Housings and Silencers: Remnant value
                Key = air_distribution_housings_and_silencers_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            return res;
        }

        double SetRadiatorsPipesElectricity(Feature building, ref CExcel exls)
        {

            String Key;
            object value;
            String cell;
            double res = 0;

            // Radiators
            #region Radiators
            #region Radiators?
            Key = change_radiators;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Radiators: Life of Product
                Key = radiators_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Radiators: Initial Investment
                Key = radiators_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Radiators: Total Installation Cost
                Key = radiators_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Radiators: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Radiators: Maintenace cost
                Key = radiators_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Radiators: Taxes amd fees cost
                Key = radiators_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Radiators: Liquidation cost
                Key = radiators_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Radiators: Remnant value
                Key = radiators_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Piping System Copper
            #region Change Piping System Copper
            #region Change Piping System Copper?
            Key = change_piping_copper;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Piping System Copper: Life of Product
                Key = piping_copper_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Copper: Initial Investment
                Key = piping_copper_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Copper: Total Installation Cost
                Key = piping_copper_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Copper: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Piping System Copper: Maintenace cost
                Key = piping_copper_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Copper: Taxes amd fees cost
                Key = piping_copper_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Copper: Liquidation cost
                Key = piping_copper_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Copper: Remnant value
                Key = piping_copper_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Piping System PEX
            #region Change Piping System PEX
            #region Change Piping System PEX?
            Key = change_piping_pex;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Piping System PEX: Life of Product
                Key = piping_pex_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PEX: Initial Investment
                Key = piping_pex_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PEX: Total Installation Cost
                Key = piping_pex_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PEX: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Piping System PEX: Maintenace cost
                Key = piping_pex_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PEX: Taxes amd fees cost
                Key = piping_pex_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PEX: Liquidation cost
                Key = piping_pex_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PEX: Remnant value
                Key = piping_pex_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Piping System PP
            #region Change Piping System PP
            #region Change Piping System PP?
            Key = change_piping_pp;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Piping System PP: Life of Product
                Key = piping_pp_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PP: Initial Investment
                Key = piping_pp_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PP: Total Installation Cost
                Key = piping_pp_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PP: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Piping System PP: Maintenace cost
                Key = piping_pp_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PP: Taxes amd fees cost
                Key = piping_pp_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PP: Liquidation cost
                Key = piping_pp_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System PP: Remnant value
                Key = piping_pp_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Piping System Cast Iron
            #region Change Piping System Cast Iron
            #region Change Piping System Cast Iron?
            Key = change_piping_cast_iron;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Piping System Cast Iron: Life of Product
                Key = piping_cast_iron_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Cast Iron: Initial Investment
                Key = piping_cast_iron_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Cast Iron: Total Installation Cost
                Key = piping_cast_iron_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Cast Iron: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Piping System Cast Iron: Maintenace cost
                Key = piping_cast_iron_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Cast Iron: Taxes amd fees cost
                Key = piping_cast_iron_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Cast Iron: Liquidation cost
                Key = piping_cast_iron_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Cast Iron: Remnant value
                Key = piping_cast_iron_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Piping System Galvanized Steel
            #region Change Piping System Galvanized Steel
            #region Change Piping System Galvanized Steel?
            Key = change_piping_galvanized_steel;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Piping System Galvanized Steel: Life of Product
                Key = piping_galvanized_steel_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Galvanized Steel: Initial Investment
                Key = piping_galvanized_steel_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Galvanized Steel: Total Installation Cost
                Key = piping_galvanized_steel_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Galvanized Steel: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Piping System Galvanized Steel: Maintenace cost
                Key = piping_galvanized_steel_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Galvanized Steel: Taxes amd fees cost
                Key = piping_galvanized_steel_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Galvanized Steel: Liquidation cost
                Key = piping_galvanized_steel_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Galvanized Steel: Remnant value
                Key = piping_galvanized_steel_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Piping System Relining
            #region Change Piping System Relining
            #region Change Piping System Relining?
            Key = change_piping_relining;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Piping System Relining: Life of Product
                Key = piping_relining_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Relining: Initial Investment
                Key = piping_relining_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Relining: Total Installation Cost
                Key = piping_relining_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Relining: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Piping System Relining: Maintenace cost
                Key = piping_relining_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Relining: Taxes amd fees cost
                Key = piping_relining_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Relining: Liquidation cost
                Key = piping_relining_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Piping System Relining: Remnant value
                Key = piping_relining_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            // Electrical Wiring
            #region Change Electrical Wiring
            #region Change Electrical Wiring?
            Key = change_electrical_wiring;
            value = (bool)building.properties[Key];
            #endregion
            if ((bool)value)
            {
                #region Electrical Wiring: Life of Product
                Key = electrical_wiring_life_of_product;
                cell = "E7";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electrical Wiring: Initial Investment
                Key = electrical_wiring_initial_investment;
                cell = "E12";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electrical Wiring: Total Installation Cost
                Key = electrical_wiring_installation_cost;
                cell = "E13";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electrical Wiring: Operating Cost
                #region Heat
                cell = "E6";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E7";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Water
                cell = "E11";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E12";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electric
                cell = "E16";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E17";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Natural gas
                cell = "E21";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));

                cell = "E22";
                value = 0.0;
                if (!exls.SetCellValue("Operating costs", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion
                #endregion

                #region Electrical Wiring: Maintenace cost
                Key = electrical_wiring_maintenance_cost;
                cell = "E21";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electrical Wiring: Taxes amd fees cost
                Key = electrical_wiring_taxes_fees_cost;
                cell = "E25";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electrical Wiring: Liquidation cost
                Key = electrical_wiring_liquidation_cost;
                cell = "E26";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                #region Electrical Wiring: Remnant value
                Key = electrical_wiring_remnant_value;
                cell = "E28";
                value = Convert.ToDouble(building.properties[Key]);
                if (!exls.SetCellValue("LCC", cell, value))
                    throw new Exception(String.Format("Could not set cell {} to value {2}", cell, value));
                #endregion

                res += Convert.ToDouble(exls.GetCellValue("LCC", result_cell));
            }
            #endregion

            return res;
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

            if (kpiId != kpi_lcc)
            {
                throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
            }

            #region Set Common Properties
            String Key;
            object value = 0;

            #region Quantity (one component at a time)
            Set(sheet: "LCC", cell: "E6", value: 1, exls: ref exls);
            #endregion

            #region Discount rate
            Key = discount_rate;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "LCC", cell: "E8", value: value, exls: ref exls);
            #endregion

            #region Heat cost
            Key = heat_cost;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "Operating costs", cell: "E8", value: value, exls: ref exls);
            #endregion

            #region Water cost
            Key = water_cost;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "Operating costs", cell: "E13", value: value, exls: ref exls);
            #endregion

            #region Electric cost
            Key = electric_cost;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "Operating costs", cell: "E18", value: value, exls: ref exls);
            #endregion

            #region Natural gas cost
            Key = natural_gas_cost;
            value = Convert.ToDouble(((Number)commonProperties[Key]).GetValue());
            Set(sheet: "Operating costs", cell: "E23", value: value, exls: ref exls);
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
                    kpi += SetInputDataOneBuilding(building, ref exls);
                }

            }

            switch (kpiId)
            {
                case kpi_lcc:
                    outputs.Add(new Kpi(Math.Round(kpi, 2), "Life cycle cost", "EUR"));
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
                LCC_Module_ErrorRaised(this, ex);
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


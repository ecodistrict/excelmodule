using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Yaml.Serialization;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace LCC
{
    class LCC_Module : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_lcc = "lcc";
        const string kpi_yearsToPayback = "years-to-payback"; //Todo: to dashboard
        const string kpi_totalLCC = "total-lcc"; //Todo: to dashboard
        const string sheet = "LCC-mall 2";
        const string buidingIdKey = "gml_id";

        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_lcc,                 "H4"},
            {kpi_yearsToPayback,      "J4"},
            {kpi_totalLCC,      "H4"}
        };

        Dictionary<string, string> generalCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"DiscountRateExclInflation",                           "C4"},
            {"ElectricityPriceIncrease",                            "C5"},
            {"DistrictHeatingPriceIncrease",                        "C6"},
            {"energy_price_increase_district_heating_natural_gas",  "C7"},
            {"energy_price_increase_other_fuel",                    "C8"},
            {"feed_in_tariff_price_increase",                       "C9"},
            {"LCACalcPeriod ",                                      "C10"},
            {"rent_increase",                                       "C11"},
            {"feed_in_tariff_price",                                "C12"},
            {"ElectricityPrice",                                    "C13"},
            {"natural_gas_price",                                   "C14"},
            {"DistrictHeatingPrice",                                "C15"},
            {"other_fule_price",                                    "C16"},
            {"rent_incomes",                                        "C17"}
        };

        Dictionary<string, string> buildingCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"fixed_cost_for_electricity",          "C24"},
            {"ElectricityConsumption",              "E25"},
            {"electricityProduction",               "E26"},
            {"fixed_cost_district_heating",         "C27"},       
            {"DistrictHeatingConsumption",          "E28"},
            {"fixed_cost_for_natural_gas",          "C29"},
            {"natural_gas_consumption ",            "E30"},
            {"fixed_cost_for_other_fuel",           "C31"},
            {"other_fuel_consumption ",             "E32"},
            {"operating_and_maintenance_costs",     "C34"}
        };

        Dictionary<string, string> heatingSystemCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"BuildingHeatingSystem_LifeOfProduct",                 "J24"}, 
            {"BuildingHeatingSystemHeatSource_InitialInvestement",  "N24"},
            {"BuildingHeatingSystemHeatSource_InstallationCost",    "O24"}
        };

        Dictionary<string, string> heatingPumpCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"BuildingSystemHeatPump_LifeOfProduct",      "J25"}, 
            {"BuildingSystemHeatPump_InitialInvestment",  "N25"},
            {"BuildingSystemHeatPump_InstallationCost",   "O25"}
        };

        Dictionary<string, string> heatingSystemBoreHoleCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"BuildingSystemBoreHole_LifeOfProduct",          "J26"}, 
            {"BuildingSystemBoreHole_InitialInvestment",      "N26"},
            {"BuildingSystemBoreHole_InstallationCost",       "O26"}
        };

        Dictionary<string, string> heatingSystemCirculationPumpCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"heating_CirculationPump_system_life_time",               "J27"}, 
            {"heating_CirculationPump_system_investment_cost",         "N27"},
            {"heating_CirculationPump_system_installation_cost",       "O27"}
        };

        Dictionary<string, string> buildingShellInsulationMaterial1CellMapping = new Dictionary<string, string>()
        {
            {"BuildingShellInsulationMaterial1_LifeOfProduct",          "J28"}, 
            {"BuildingShellInsulationMaterial1_InitialInvestement",     "N28"},
            {"BuildingShellInsulationMaterial1_InstallationCost",       "O28"}
        };

        Dictionary<string, string> buildingShellInsulationMaterial2CellMapping = new Dictionary<string, string>()
        {
            {"BuildingShellInsulationMaterial2_LifeOfProduct",          "J29"}, 
            {"BuildingShellInsulationMaterial2_InitialInvestement",     "N29"},
            {"BuildingShellInsulationMaterial2_InstallationCost",       "O29"}
        };

        Dictionary<string, string> buildingShellFacadeSystemCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"Building_ShellFacade_system_life_time",               "J30"}, 
            {"Building_ShellFacade_system_investment_cost",         "N30"},
            {"Building_ShellFacade_system_installation_cost",       "O30"}
        };

        Dictionary<string, string> buildingShellWindowsCellMapping = new Dictionary<string, string>()
        {
            {"BuildingShellWindows_LifeOfProduct",        "J31"}, 
            {"BuildingShellWindows_InitialInvestement",   "N31"},
            {"BuildingShellWindows_InstallationCost",     "O31"}
        };

        Dictionary<string, string> buildingShellDoorsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"Building_ShellDoors_system_life_time",               "J32"}, 
            {"Building_ShellDoors_system_investment_cost",         "N32"},
            {"Building_ShellDoors_system_installation_cost",       "O32"}
        };

        Dictionary<string, string> VentilationSystemVentilationDuctsCellMapping = new Dictionary<string, string>()
        {
            {"VentilationSystem_LifeOfVentilationDucts",      "J33"}, 
            {"VentilationSystem_InitialInvestementDucts",     "N33"},
            {"VentilationSystem_InstallationCostDucts",       "O33"}
        };

        Dictionary<string, string> VentilationSystemAirflowAssemblyCellMapping = new Dictionary<string, string>()
        {
            {"VentilationSystem_LifeOfAirflowAssembly",                 "J34"}, 
            {"VentilationSystem_InitialInvestementAirflowAssembly",     "N34"},
            {"VentilationSystem_InstallationCostAirflowAssembly",       "O34"}
        };

        Dictionary<string, string> VentilationSystemDistributionHousingsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"VentilationSystemDistributionHousings_life_time",               "J35"}, 
            {"VentilationSystemDistributionHousings_investment_cost",         "N35"},
            {"VentilationSystemDistributionHousings_installation_cost",       "O35"}
        };

        Dictionary<string, string> RadiatorsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"BuildingRadiators_LifeOfProduct",          "J36"}, 
            {"BuildingRadiators_InitialInvestement",     "N36"},
            {"BuildingRadiators_InstallationCost",       "O36"}
        };

        Dictionary<string, string> WaterTapsCellMapping = new Dictionary<string, string>()
        {
            {"WaterTaps_LifeOfProduct",               "J37"}, 
            {"WaterTaps_InitialInvestement",          "N37"},
            {"WaterTaps_InstallationCost",            "O37"}
        };

        Dictionary<string, string> PipingSystemsCopperCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"PipingSystemsCopper_LifeOfProduct",               "J38"}, 
            {"PipingSystemsCopper_InitialInvestement",  "N38"},
            {"PipingSystemsCopper_InstallationCost",    "O38"}
        };

        Dictionary<string, string> PipingSystemsPEXCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"PipingSystemsPEX_life_time",               "J39"}, 
            {"PipingSystemsPEX_investment_cost",         "N39"},
            {"PipingSystemsPEX_installation_cost",       "O39"}
        };

        Dictionary<string, string> PipingSystemsPPCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"PipingSystemsPP_life_time",               "J40"}, 
            {"PipingSystemsPP_investment_cost",         "N40"},
            {"PipingSystemsPP_installation_cost",       "O40"}
        };

        Dictionary<string, string> PipingSystemsCastIronCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"PipingSystemsCastIron_life_time",               "J41"}, 
            {"PipingSystemsCastIron_investment_cost",         "N41"},
            {"PipingSystemsCastIron_installation_cost",       "O41"}
        };

        Dictionary<string, string> PipingSystemsGalvanisedSteelCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"PipingSystemsGalvanisedSteel_life_time",               "J42"}, 
            {"PipingSystemsGalvanisedSteel_investment_cost",         "N42"},
            {"PipingSystemsGalvanisedSteel_installation_cost",       "O42"}
        };

        Dictionary<string, string> PipingSystemsReliningCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"PipingSystemsRelining_life_time",               "J43"}, 
            {"PipingSystemsRelining_investment_cost",         "N43"},
            {"PipingSystemsRelining_installation_cost",       "O43"}
        };

        Dictionary<string, string> ElectricalWiringCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"ElectricalWiring_life_time",               "J44"}, 
            {"ElectricalWiring_investment_cost",         "N44"},
            {"ElectricalWiring_installation_cost",       "O44"}
        };


        Dictionary<string, string> EnergyProductionCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"EnergyProduction_LifeOfProduct",          "J45"}, 
            {"EnergyProduction_InitialInvestement",     "N45"},
            {"EnergyProduction_InstallationCost",       "O45"}
        };

        Dictionary<string, string> BuildingCondBoilersCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"BuildingCondBoiler_LifeOfProduct",          "J46"}, 
            {"BuildingCondBoiler_InitialInvestement",     "N46"},
            {"BuildingCondBoiler_InstallationCost",       "O46"}
        };


        #endregion

        #endregion
        
        public LCC_Module()
        {
            this.useDummyDB = false;
            this.useBothVariantAndAsISForVariant = false;
            
            //List of kpis the module can calculate
            this.KpiList = kpiCellMapping.Keys.ToList();

            //Error handler
            this.ErrorRaised += CExcelModule_ErrorRaised;

            //Notification
            this.StatusMessage += CExcelModule_StatusMessage;            
        }

        private bool SetProperties(ModuleProcess process, CExcel exls, Dictionary<string, string> propertyCellMapping)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                Dictionary<string, object> CurrentData = process.CurrentData;
                try
                {
                    {

                        if (CurrentData.ContainsKey(property.Key))
                        {
	                        object value = CurrentData[property.Key];
	
	                        double val = Convert.ToDouble(value);
	                        if (val < 0)
	                        {
                                process.CalcMessage = String.Format("Property '{0}' has invalid data, only values equal or above zero is allowed; value: {1}", property.Key, val);
	                            return false;
	                        }
	
	                        Set(sheet, property.Value, value, ref exls);
                        }
                        else
                        {
                            process.CalcMessage = "";
                            return false;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    SendErrorMessage(message: String.Format(ex.Message + "\t key = {0}, isCurrentDataMissing = {1}", property.Key, CurrentData == null), sourceFunction: "SetProperties", exception: ex);
                    throw ex;
                }
            }

            return true;
        }

        private bool SetProperties(Dictionary<string, object> buildingData, CExcel exls, Dictionary<string, string> propertyCellMapping, out bool changesMade)
        {
            changesMade = false;
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                try
                {
                    if (buildingData.ContainsKey(property.Key))
                    {
                        object value = buildingData[property.Key];
                        Set(sheet, property.Value, value, ref exls);
                        changesMade = true;
                    }
                    //else
                    //{
                    //    Set(sheet, property.Value, 0, ref exls);
                    //    //TODO
                    //}

                }
                catch (System.Exception ex)
                {
                    SendErrorMessage(message: String.Format(ex.Message + "\t key = {0}", property.Key), sourceFunction: "SetProperties", exception: ex);
                    throw ex;
                }
            }

            return true;
        }

        private bool GetProperties(CExcel exls, Dictionary<string, string> cellMapping, ref Dictionary<string, object> buildingDefaultValues)
        {
            foreach (KeyValuePair<string,string> property in cellMapping)
            {
                try
                {
                    buildingDefaultValues.Add(property.Key,exls.GetCellValue(sheet,property.Value));
                }
                catch (Exception ex)
                {
                    SendErrorMessage(message: String.Format(ex.Message + "\t key = {0}", property.Key), sourceFunction: "GetProperties", exception: ex);
                    return false;
                }
            }
            return true;
        }

        protected override bool CalculateKpi(ModuleProcess process, CExcel exls, out Ecodistrict.Messaging.Data.Output output, out Ecodistrict.Messaging.Data.OutputDetailed outputDetailed)
        {
            try
            {
                output = null;
                outputDetailed = null;

                //Check and prepare data
                //Dictionary<string, object> district_data;
                //GeoValue buildingsAsIS;
                if (!CheckAndReportDistrictProp(process,process.CurrentData, "Buildings"))
                    return false;

                string buildingsData = Newtonsoft.Json.JsonConvert.SerializeObject(process.CurrentData["Buildings"]);  //TODO error if process.CurrentData["Buildings"] == "{}"
                List<Dictionary<string, object>> buildings = Newtonsoft.Json.JsonConvert.DeserializeObject(buildingsData, typeof(List<Dictionary<string, object>>)) as List<Dictionary<string, object>>;

                //buildings = process.CurrentData["Buildings"] as GeoValue;
                //buildings = Ecodistrict.Messaging.DeserializeData<GeoValue>.JsonString(process.CurrentData["Buildings"] as string);
                //bool perHeatedArea;
                //if (!CheckAndPrepareData(process, out district_data, out buildingsAsIS, out buildingsVariant, out perHeatedArea))
                //    return false;

                //Set common properties
                if (!SetProperties(process, exls, generalCellMapping))
                    return false;
                
                //Calculate kpi
                outputDetailed = new Ecodistrict.Messaging.Data.OutputDetailed(process.KpiId);
                double kpiValue = 0;
                int noRenovatedBuildings = 0;

                //NEW_CODE: Start with getting all default (Inital start) values from the excel Sheet for buildingdata
                Dictionary<string,object> buildingDefaultValues=new Dictionary<string, object>();
                if(buildings != null && buildings.Count>0)
                    if (!GetBuildingDefaultValues(exls, out buildingDefaultValues))
                        return false;

                foreach (Dictionary<string,object> buildingData in buildings)
                {
                    double kpiValuei;
                    bool changesMade;
                    if (!SetInputDataOneBuilding(buildingData, exls, out changesMade))
                        return false;

                    kpiValuei = Convert.ToDouble(exls.GetCellValue(sheet, kpiCellMapping[process.KpiId]));
                    
                    if (changesMade)
                        ++noRenovatedBuildings;

                    //NEW_CODE: Reset all used building values
                    if (buildingDefaultValues != null && 
                        !SetInputDataOneBuilding(buildingDefaultValues, exls, out changesMade))
                        return false;
                    
                    kpiValue += kpiValuei;
                    outputDetailed.KpiValueList.Add(new Ecodistrict.Messaging.Data.GeoObject("building", buildingData[buidingIdKey] as string, process.KpiId, kpiValuei));
                }

                if (noRenovatedBuildings > 0 && (process.KpiId != kpi_totalLCC))
                    kpiValue /= Convert.ToDouble(noRenovatedBuildings);

                output = new Ecodistrict.Messaging.Data.Output(process.KpiId, Math.Round(kpiValue, 1));

                return true;
            }
            catch (System.Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "CalculateKpi", exception: ex);
                throw ex;
            }
        }

        private bool GetBuildingDefaultValues(CExcel exls, out Dictionary<string, object> buildingDefaultValues)
        {
            buildingDefaultValues=new Dictionary<string, object>();
            try
            {
                #region Get data

                if (!GetProperties(exls, buildingCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, buildingShellInsulationMaterial1CellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, buildingShellWindowsCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, VentilationSystemVentilationDuctsCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, VentilationSystemAirflowAssemblyCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, WaterTapsCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, PipingSystemsCopperCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, PipingSystemsPEXCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, PipingSystemsPPCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, PipingSystemsCastIronCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, PipingSystemsGalvanisedSteelCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, PipingSystemsReliningCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, ElectricalWiringCellMapping, ref buildingDefaultValues))
                    return false;
                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        bool SetInputDataOneBuilding(Dictionary<string, object> buildingData, CExcel exls, out bool changesMade)
        {
            changesMade = false;
            bool changesMade_i = false;

            try
            {
                #region Set Data
                if (!SetProperties(buildingData, exls, buildingCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, buildingShellInsulationMaterial1CellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, buildingShellWindowsCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, VentilationSystemVentilationDuctsCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, VentilationSystemAirflowAssemblyCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, WaterTapsCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, PipingSystemsCopperCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, PipingSystemsPEXCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, PipingSystemsPPCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, PipingSystemsCastIronCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, PipingSystemsGalvanisedSteelCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;
                
                if (!SetProperties(buildingData, exls, PipingSystemsReliningCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, ElectricalWiringCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;
                
                #endregion

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }

        private void Set(string sheet, string cell, object value, ref CExcel exls)
        {
            if (!exls.SetCellValue(sheet, cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {2} in sheet {3}", cell, value, sheet));
        }

        bool GetSetBool(ref Dictionary<string, object> properties, string property)
        {
            if (!properties.ContainsKey(property))
                properties.Add(property, false);

            if (properties[property] is bool)
                return (bool)properties[property];

            return false;
        }
        
    }
}


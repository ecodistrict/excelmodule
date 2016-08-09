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
        const string sheet = "LCC-mall 2";
        const string buidingIdKey = "gml_id";

        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_lcc,                 "H4"}
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
            //{"heating_system_time_for_investment",     "G20"},
            {"BuildingHeatingSystem_LifeOfProduct",                 "J24"}, 
            {"BuildingHeatingSystemHeatSource_InitialInvestement",  "N24"},
            {"BuildingHeatingSystemHeatSource_InstallationCost",    "O24"}
        };

        Dictionary<string, string> heatingSystemBoreHoleCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J25"}, 
            {"heating_system_investment_cost",         "N25"},
            {"heating_system_installation_cost",       "O25"}
        };

        Dictionary<string, string> heatingSystemCirculationPumpCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J26"}, 
            {"heating_system_investment_cost",         "N26"},
            {"heating_system_installation_cost",       "O26"}
        };

        Dictionary<string, string> buildingShellInsulationMaterial1CellMapping = new Dictionary<string, string>()
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"BuildingShellInsulationMaterial1_LifeOfProduct",          "J27"}, 
            {"BuildingShellInsulationMaterial1_InitialInvestement",     "N27"},
            {"BuildingShellInsulationMaterial1_InstallationCost",       "O27"}
        };

        Dictionary<string, string> buildingShellInsulationMaterial2CellMapping = new Dictionary<string, string>()
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"BuildingShellInsulationMaterial2_LifeOfProduct",               "J28"}, 
            {"BuildingShellInsulationMaterial2_InitialInvestement",         "N28"},
            {"BuildingShellInsulationMaterial2_InstallationCost",       "O28"}
        };

        Dictionary<string, string> buildingShellFacadeSystemCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J29"}, 
            {"heating_system_investment_cost",         "N29"},
            {"heating_system_installation_cost",       "O29"}
        };

        Dictionary<string, string> buildingShellWindowsCellMapping = new Dictionary<string, string>()
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"BuildingShellWindows_LifeOfProduct",        "J30"}, 
            {"BuildingShellWindows_InitialInvestement",   "N30"},
            {"BuildingShellWindows_InstallationCost",     "O30"}
        };

        Dictionary<string, string> buildingShellDoorsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J31"}, 
            {"heating_system_investment_cost",         "N31"},
            {"heating_system_installation_cost",       "O31"}
        };

        Dictionary<string, string> VentilationSystemVentilationDuctsCellMapping = new Dictionary<string, string>()
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"VentilationSystem_LifeOfVentilationDucts",      "J32"}, 
            {"VentilationSystem_InitialInvestementDucts",     "N32"},
            {"VentilationSystem_InstallationCostDucts",       "O32"}
        };

        Dictionary<string, string> VentilationSystemAirflowAssemblyCellMapping = new Dictionary<string, string>()
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"VentilationSystem_LifeOfAirflowAssembly",                 "J33"}, 
            {"VentilationSystem_InitialInvestementAirflowAssembly",     "N33"},
            {"VentilationSystem_InstallationCostAirflowAssembly",       "O33"}
        };

        Dictionary<string, string> VentilationSystemDistributionHousingsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J34"}, 
            {"heating_system_investment_cost",         "N34"},
            {"heating_system_installation_cost",       "O34"}
        };

        Dictionary<string, string> RadiatorsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"BuildingRadiators_LifeOfProduct",               "J35"}, 
            {"BuildingRadiators_InitialInvestement",         "N35"},
            {"BuildingRadiators_InstallationCost",       "O35"}
        };

        Dictionary<string, string> WaterTapsCellMapping = new Dictionary<string, string>()
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"WaterTaps_LifeOfProduct",               "J36"}, 
            {"WaterTaps_InitialInvestement",         "N36"},
            {"WaterTaps_InstallationCost",       "O36"}
        };

        Dictionary<string, string> PipingSystemsCopperCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"PipingSystemsCopper_LifeOfProduct",               "J37"}, 
            {"BuildingPipingSystemsCopper_InitialInvestement",         "N37"},
            {"BuildingPipingSystemsCopper_InstallationCost",       "O37"}
        };

        Dictionary<string, string> PipingSystemsPEXCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J38"}, 
            {"heating_system_investment_cost",         "N38"},
            {"heating_system_installation_cost",       "O38"}
        };

        Dictionary<string, string> PipingSystemsPPCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J39"}, 
            {"heating_system_investment_cost",         "N39"},
            {"heating_system_installation_cost",       "O39"}
        };

        Dictionary<string, string> PipingSystemsCastIronCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J40"}, 
            {"heating_system_investment_cost",         "N40"},
            {"heating_system_installation_cost",       "O40"}
        };

        Dictionary<string, string> PipingSystemsGalvanisedSteelCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J41"}, 
            {"heating_system_investment_cost",         "N41"},
            {"heating_system_installation_cost",       "O41"}
        };

        Dictionary<string, string> PipingSystemsReliningCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J42"}, 
            {"heating_system_investment_cost",         "N42"},
            {"heating_system_installation_cost",       "O42"}
        };

        Dictionary<string, string> ElectricalWiringCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"heating_system_life_time",               "J43"}, 
            {"heating_system_investment_cost",         "N43"},
            {"heating_system_installation_cost",       "O43"}
        };

        Dictionary<string, string> EnergyProductionCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            //{"heating_system_time_for_investment",     "G20"},
            {"EnergyProduction_LifeOfProduct",               "J44"}, 
            {"EnergyProduction_InitialInvestement",         "N44"},
            {"EnergyProduction_InstallationCost",       "O44"}
        };

        #endregion

        #endregion
        
        public LCC_Module()
        {
            this.useDummyDB = false;
            this.useBothVariantAndAsISForVariant = false;
            
            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_lcc };

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

                string buildingsData = Newtonsoft.Json.JsonConvert.SerializeObject(process.CurrentData["Buildings"]);
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
                foreach (Dictionary<string,object> buildingData in buildings)
                {
                    double kpiValuei;
                    bool changesMade;
                    if (!SetInputDataOneBuilding(buildingData, exls, out changesMade))
                        return false;

                    kpiValuei = Convert.ToDouble(exls.GetCellValue(sheet, kpiCellMapping[process.KpiId]));

                    if (changesMade)
                        ++noRenovatedBuildings;
                    
                    kpiValue += kpiValuei;
                    outputDetailed.KpiValueList.Add(new Ecodistrict.Messaging.Data.GeoObject("building", buildingData[buidingIdKey] as string, process.KpiId, kpiValuei));
                }

                if (noRenovatedBuildings > 0)
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


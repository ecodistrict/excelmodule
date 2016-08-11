using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Yaml.Serialization;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace Affordability
{
    class Affordability_Module : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_mdi = "minimum-disposable-income";
        const string sheet = "Customised input";
        const string sheetOutput = "Output";
        const string buidingIdKey = "gml_id";

        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_mdi,                 "C22"}
        };

        private Dictionary<string, string> generalCellMapping = new Dictionary<string, string>()
        {
            {"maximum_housing_costs_as_a_percentage_of_disposable_income", "C17"}
        };
        Dictionary<string, string> generalBuildingCellMapping = new Dictionary<string, string>()
        {
            {"size_of_household",                                           "C5"},
            {"size_of_dwelling",                                            "C6"},
            {"rent_or_cost_of_housing_per_month",                           "C10"}
        };

        #region Houshold Electricity
        Dictionary<string, string> householdElectricityCellMapping = new Dictionary<string, string>()
        {
            {include_household_electricity.Key,     include_household_electricity.Value},
            {useTypicalHouseholdElectricity.Key,    useTypicalHouseholdElectricity.Value},
            {energy_price_houshold_electricity.Key, energy_price_houshold_electricity.Value}
        };


        Checkbox cb_include_household_electricity = new Checkbox("Household electricity payed separately");

        static KeyValuePair<string, string> include_household_electricity =
            new KeyValuePair<string, string>("include_household_electricity", "C12");

        static KeyValuePair<string, string> energy_price_houshold_electricity =
            new KeyValuePair<string, string>("energy_price_houshold_electricity", "F23");

        Select householdElectricitySelect = new Select("Use typical calculation of household electricity (based on household size and heated floor area)?", householdElectricityOpt, null, householdElectricityOpt.First());

        static Options householdElectricityOpt = new Options()
        {
            new Option("no - enter manually below","no - enter manually below"),
            new Option("yes, energy efficient white goods, appliances and lights are used","yes, energy efficient white goods, appliances and lights are used"),
            new Option("yes, conventional appliances are used","yes, conventional appliances are used")
        };

        static KeyValuePair<string, string> useTypicalHouseholdElectricity =
            new KeyValuePair<string, string>("use_typical_calculation_for_household_electricity", "C22");
        KeyValuePair<string, string> manualHouseholdElectricity =
            new KeyValuePair<string, string>("annual_energy_consumption_for_household_electricity", "C23");
        #endregion

        #region DHW

        Checkbox cb_include_domestic_hot_water = new Checkbox("Domestic hot water payed separately");

        KeyValuePair<string, string> include_domestic_hot_water =
            new KeyValuePair<string, string>("include_domestic_hot_water", "C13");

        KeyValuePair<string, string> energy_price_domestic_hot_water =
            new KeyValuePair<string, string>("energy_price__domestic_hot_water", "F27");

        Select dHWSelect = new Select("Use typical calculation of domestic hot water (based on household size)?", dHWOpt, null, dHWOpt.First());

        static Options dHWOpt = new Options()
        {
            new Option("no - enter manually below","no - enter manually below"),
            new Option("yes, water saving/energy efficient taps are used","yes, water saving/energy efficient taps are used"),
            new Option("yes, conventional taps are used","yes, conventional taps are used")
        };

        static KeyValuePair<string, string> useTypicalDHW =
            new KeyValuePair<string, string>("use_typical_calculation_for_domestic_hot_water", "C26");
        KeyValuePair<string, string> manualDHW =
            new KeyValuePair<string, string>("annual_energy_consumption_for_domestic_hot_water", "C27");
        #endregion

        #region Cooling
        Checkbox cb_include_cooling = new Checkbox("Cooling payed separately");

        KeyValuePair<string, string> include_cooling =
            new KeyValuePair<string, string>("include_cooling", "C14");

        KeyValuePair<string, string> energy_price_cooling =
            new KeyValuePair<string, string>("energy_price_cooling", "F31");

        KeyValuePair<string, string> manualCooling =
            new KeyValuePair<string, string>("annual_energy_consumption_for_cooling", "C31");
        #endregion

        #region Space Heating

        Checkbox cb_include_space_heating = new Checkbox("Space heating payed separately");

        KeyValuePair<string, string> include_space_heating =
            new KeyValuePair<string, string>("include_space_heating", "C15");

        KeyValuePair<string, string> energy_price_space_heating =
            new KeyValuePair<string, string>("energy_price_space_heating", "F35");

        KeyValuePair<string, string> manualSpaceHeating =
            new KeyValuePair<string, string>("annual_energy_consumption_for_space_heating", "C35");
        #endregion

        #endregion

        #endregion

        public Affordability_Module()
        {
            this.useBothVariantAndAsISForVariant = false;
            
            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_mdi };

            //Error handler
            this.ErrorRaised += CExcelModule_ErrorRaised;

            //Notification
            this.StatusMessage += CExcelModule_StatusMessage;            
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
                if (!CheckAndReportDistrictProp(process, process.CurrentData, "Buildings"))
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
                foreach (Dictionary<string, object> buildingData in buildings)
                {
                    double kpiValuei;
                    bool changesMade;
                    if (!SetInputDataOneBuilding(process,buildingData, exls, out changesMade))
                        return false;

                    kpiValuei = 100 * Convert.ToDouble(exls.GetCellValue(sheetOutput, kpiCellMapping[process.KpiId]));
                    
                    //Todo: Reset all used building values

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

        private bool SetInputDataOneBuilding(ModuleProcess process, Dictionary<string, object> buildingData, CExcel exls, out bool changesMade)
        {
            changesMade = false;

            if (!Set(process, buildingData, generalBuildingCellMapping, ref exls))
                return false;

            #region Household Electricity
            if (buildingData.ContainsKey(include_household_electricity.Key))
            {
                if (!Set(process, buildingData, include_household_electricity, ref exls))
                    return false;

                if ((bool)buildingData[include_household_electricity.Key])
                {
                    if (!Set(process, buildingData, energy_price_houshold_electricity, ref exls))
                        return false;

                    if (!Set(process, buildingData, useTypicalHouseholdElectricity, ref exls))
                        return false;

                    if ((string)buildingData[useTypicalHouseholdElectricity.Key] == householdElectricityOpt.First().value)
                    {
                        if (!Set(process, buildingData, manualHouseholdElectricity, ref exls))
                            return false;
                    }
                }
                changesMade = true;
            }
            #endregion

            #region DHW
            if (buildingData.ContainsKey(include_domestic_hot_water.Key))
            {
                if (!Set(process, buildingData, energy_price_domestic_hot_water, ref exls))
                    return false;

                if (!Set(process, buildingData, include_domestic_hot_water, ref exls))
                    return false;

                if ((bool)buildingData[include_domestic_hot_water.Key])
                {
                    if (!Set(process, buildingData, useTypicalDHW, ref exls))
                        return false;

                    if ((string)buildingData[useTypicalDHW.Key] == dHWOpt.First().value)
                    {
                        if (!Set(process, buildingData, manualDHW, ref exls))
                            return false;
                    }
                }
                changesMade = true;
            }
            #endregion

            #region Cooling
            if (buildingData.ContainsKey(include_cooling.Key))
            {
                if (!Set(process, buildingData, energy_price_cooling, ref exls))
                    return false;

                if (!Set(process, buildingData, include_cooling, ref exls))
                    return false;

                if ((bool)buildingData[include_cooling.Key])
                {
                    if (!Set(process, buildingData, manualCooling, ref exls))
                        return false;
                }
                changesMade = true;
            }
            #endregion

            #region Space Heating
            if (buildingData.ContainsKey(include_space_heating.Key))
            {
                if (!Set(process, buildingData, energy_price_space_heating, ref exls))
                    return false;

                if (!Set(process, buildingData, include_space_heating, ref exls))
                    return false;

                if ((bool)buildingData[include_space_heating.Key])
                {
                    if (!Set(process, buildingData, manualSpaceHeating, ref exls))
                        return false;
                }
                changesMade = true;
            }
            #endregion


            return true;

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

        private bool Set(ModuleProcess process, Dictionary<string, object> buildingData, Dictionary<string, string> propertyCellMappings, ref CExcel exls)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMappings)
            {
                if (!Set(process, buildingData, property, ref exls))
                    return false;
            }

            return true;
        }

        private bool Set(ModuleProcess process, Dictionary<string, object> buildingData, KeyValuePair<string, string> propertyCellMapping, ref CExcel exls)
        {
            if (!CheckAndReportBuildingProp(process, buildingData, propertyCellMapping.Key))
                return false;

            Set(sheet, propertyCellMapping.Value, buildingData[propertyCellMapping.Key], ref exls);

            return true;
        }



        //Old system
        void SetIspec(ref InputSpecification iSpec, Dictionary<string, string> propertyCellMapping)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                if (!iSpec.ContainsKey(property.Key))
                    iSpec.Add(property.Key, new Number(property.Key));
            }
        }
        
        void SetInp(ref NonAtomic input, Dictionary<string, string> propertyCellMapping)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                if (!input.ContainsKey(property.Key))
                    input.Add(property.Key, new Number(property.Key));
            }
        }

        void SetInp(ref NonAtomic input, KeyValuePair<string, string> property)
        {
                if (!input.ContainsKey(property.Key))
                    input.Add(property.Key, new Number(property.Key));
           
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

        private bool Set(Feature building, Dictionary<string, string> propertyCellMappings, ref CExcel exls)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMappings)
            {
                if(!Set(building, property, ref exls))
                    return false;
            }

            return true;
        }

        private bool Set(Feature building, KeyValuePair<string, string> propertyCellMapping, ref CExcel exls)
        {
            //if (!CheckAndReportBuildingProp(building, propertyCellMapping.Key))
            //        return false;

            Set(sheet, propertyCellMapping.Value, building.properties[propertyCellMapping.Key], ref exls);                      

            return true;
        }

        //protected override bool CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls, out Ecodistrict.Messaging.Output.Outputs outputs)
        //{
        //    outputs = null;
        //    double kpi = 0;

        //    if (!indata.ContainsKey("buildings"))
        //        return false;

        //    GeoJson buildingProperties = indata["buildings"] as GeoJson;

        //    if (buildingProperties == null)
        //        return false;

        //    int nrBuildings = 0;
        //    foreach (Feature building in buildingProperties.value.features)
        //    {
        //        Set(building, generalBuildingCellMapping, ref exls);

        //        #region Household Electricity
        //        if (building.properties.ContainsKey(include_household_electricity.Key))
        //        {
        //            if (!Set(building, include_household_electricity, ref exls))
        //                return false;

        //            if ((bool)building.properties[include_household_electricity.Key])
        //            {
        //                if (!Set(building, energy_price_houshold_electricity, ref exls))
        //                    return false;

        //                if (!Set(building, useTypicalHouseholdElectricity, ref exls))
        //                    return false;

        //                if ((string)building.properties[useTypicalHouseholdElectricity.Key] == householdElectricityOpt.First().value)
        //                {
        //                    if (!Set(building, manualHouseholdElectricity, ref exls))
        //                        return false;
        //                }
        //            }
        //        }
        //        #endregion

        //        #region DHW
        //        if (building.properties.ContainsKey(include_domestic_hot_water.Key))
        //        {
        //            if (!Set(building, energy_price_domestic_hot_water, ref exls))
        //                return false;

        //            if (!Set(building, include_domestic_hot_water, ref exls))
        //                return false;

        //            if ((bool)building.properties[include_domestic_hot_water.Key])
        //            {
        //                if (!Set(building, useTypicalDHW, ref exls))
        //                    return false;

        //                if ((string)building.properties[useTypicalDHW.Key] == dHWOpt.First().value)
        //                {
        //                    if (!Set(building, manualDHW, ref exls))
        //                        return false;
        //                }
        //            }
        //        }
        //        #endregion


        //        #region Cooling
        //        if (building.properties.ContainsKey(include_cooling.Key))
        //        {
        //            if (!Set(building, energy_price_cooling, ref exls))
        //                return false;

        //            if (!Set(building, include_cooling, ref exls))
        //                return false;

        //            if ((bool)building.properties[include_cooling.Key])
        //            {
        //               if (!Set(building, manualCooling, ref exls))
        //                    return false;
        //            }
        //        }
        //        #endregion

        //        #region Space Heating
        //        if (building.properties.ContainsKey(include_space_heating.Key))
        //        {
        //            if (!Set(building, energy_price_space_heating, ref exls))
        //                return false;

        //            if (!Set(building, include_space_heating, ref exls))
        //                return false;

        //            if ((bool)building.properties[include_space_heating.Key])
        //            {
        //                if (!Set(building, manualSpaceHeating, ref exls))
        //                    return false;
        //            }
        //        }
        //        #endregion

        //        #region Result

        //        double resi = 100*Convert.ToDouble(exls.GetCellValue(sheetOutput, kpiCellMapping[kpiId]));
        //        kpi += resi;

        //        if (!building.properties.ContainsKey("kpiValue"))
        //            building.properties.Add("kpiValue", resi);
        //        else
        //            building.properties["kpiValue"] = Math.Round(resi, 0);

        //        #endregion

        //        ++nrBuildings;
        //    }
        //    kpi = kpi / (double)nrBuildings;


        //    outputs = new Ecodistrict.Messaging.Output.Outputs();

        //    switch (kpiId)
        //    {
        //        case kpi_mdi:
        //            outputs.Add(new Ecodistrict.Messaging.Output.Kpi(Math.Round(kpi, 0), "Minimum disposable income", "EUR / month"));
        //            break;
        //        default:
        //            throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
        //    }

        //    Ecodistrict.Messaging.Output.GeoJson buildingsProps = new Ecodistrict.Messaging.Output.GeoJson(buildingProperties);
        //    outputs.Add(buildingsProps);

        //    return true;
        //}

    }
}


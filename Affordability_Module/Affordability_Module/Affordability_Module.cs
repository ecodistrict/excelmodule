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
        const string sheet = "Input";
        const string sheetOutput = "Output";
        const string buidingIdKey = "gml_id";

        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_mdi,                 "C22"}
        };

        private Dictionary<string, string> generalCellMapping = new Dictionary<string, string>()
        {
            {"MaxCostShare", "B18"}
        };

        Dictionary<string, string> generalBuildingCellMapping = new Dictionary<string, string>()
        {
            {"NoOfPeopleInHousehold",   "B4"},
            {"HeatedFloorArea",         "B6"},
            {"RentPerMonth",            "B10"}
        };

        
        #region paidSeparatelyChoise

        private Dictionary<string, string> paidSeparatelyChoiseCellMapping = new Dictionary<string, string>()
            {
                {"ElectricityPaidSeparately",   "C13"}, //True/False or 0/1
                {"HotWaterPaidSerparately",     "C14"},
                {"CoolingPaidSeparately",       "C15"},
                {"SpaceHeatingPaidSeparately",  "C16"}
            };
        #endregion
        #region paidSeparatelyValuesCellMapping
        Dictionary<string, string> householdElectricityCellMapping = new Dictionary<string, string>()
        {
            {"ElectricityCalculationChoice",     "C23"},
            {"HouseholdElectricityConsumption",  "B24"},
            {"HouseHoldElectricityPrice",        "D24"}
        };

        Dictionary<string, string> householdHotWaterCellMapping = new Dictionary<string, string>()
        {
            {"HouseholdHotWaterCalculationChoice",    "C27"},
            {"HouseholdHotWaterEnergyConsumption",    "B28"},
            {"HouseholdHotWaterEnergyPrice",          "D28"}
        };

                Dictionary<string, string> householdCoolingCellMapping = new Dictionary<string, string>()
        {
            {"HouseholdCoolingEnergyConsumption",      "B32"},
            {"HouseholdCoolingEnergyConsumptionPSqM",  "B32"}, //Has to be multiplied with Dwellingsize (B6) before used in cell
            {"HouseholdCoolingEnergyPrice",            "D32"}
        };
        Dictionary<string, string> householdSpaceHeatingCellMapping = new Dictionary<string, string>()
        {
            {"HouseholdSpaceHeatingEnergyConsumption",       "B35"},
            {"HouseholdSpaceHeatingEnergyConsumptionPSqM",   "B35"}, //Has to be multiplied with Dwellingsize (B6) before used in cell
            {"HouseholdSpaceHeatingEnergyPrice",             "D35"}
        };
        #endregion

        #endregion

  

        #endregion


        public Affordability_Module()
        {
            this.useDummyDB = false;
            this.useBothVariantAndAsISForVariant = false;
            
            //List of kpis the module can calculate
            //this.KpiList = new List<string> { kpi_mdi };
            this.KpiList = kpiCellMapping.Keys.ToList();

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

                //NEW_CODE: Start with getting all default (Inital start) values from the excel Sheet for buildingdata
                Dictionary<string, object> buildingDefaultValues = new Dictionary<string, object>();
                if (buildings != null && buildings.Count > 0)
                    if (!GetBuildingDefaultValues(exls, out buildingDefaultValues))
                        return false;

                foreach (Dictionary<string, object> buildingData in buildings)
                {
                    double kpiValuei;
                    bool changesMade;
                    if (!SetInputDataOneBuilding(process,buildingData, exls, out changesMade))
                        return false;

                    kpiValuei = 100 * Convert.ToDouble(exls.GetCellValue(sheetOutput, kpiCellMapping[process.KpiId]));
                    
                    if (changesMade)
                        ++noRenovatedBuildings;

                    //NEW_CODE: Reset all used building values
                    if (!SetInputDataOneBuilding(process, buildingDefaultValues, exls, out changesMade))
                        return false;

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

        private bool GetBuildingDefaultValues(CExcel exls, out Dictionary<string, object> DefaultValues)
        {
            DefaultValues = new Dictionary<string, object>();
            try
            {
                #region Get data

                if (!GetProperties(exls, generalBuildingCellMapping, ref DefaultValues))
                    return false;
                if (!GetProperties(exls, paidSeparatelyChoiseCellMapping, ref DefaultValues))
                    return false;
                if (!GetProperties(exls, householdElectricityCellMapping, ref DefaultValues))
                    return false;
                if (!GetProperties(exls, householdHotWaterCellMapping, ref DefaultValues))
                    return false;
                if (!GetProperties(exls, householdCoolingCellMapping, ref DefaultValues))
                    return false;
                if (!GetProperties(exls, householdSpaceHeatingCellMapping, ref DefaultValues))
                    return false;

                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }


        private bool SetInputDataOneBuilding(ModuleProcess process, Dictionary<string, object> buildingData, CExcel exls, out bool changesMade)
        {
            changesMade = false;
            bool changesMade_i = false;

            try
            {
                #region Set data

                if (!SetProperties(buildingData, exls, generalBuildingCellMapping, out changesMade_i))
                    return false;
                if (!SetProperties(buildingData, exls, paidSeparatelyChoiseCellMapping, out changesMade_i))
                    return false;
                if (!SetProperties(buildingData, exls, householdElectricityCellMapping, out changesMade_i))
                    return false;
                if (!SetProperties(buildingData, exls, householdHotWaterCellMapping, out changesMade_i))
                    return false;
                if (!SetProperties(buildingData, exls, householdCoolingCellMapping, out changesMade_i))
                    return false;
                if (!SetProperties(buildingData, exls, householdSpaceHeatingCellMapping, out changesMade_i))
                    return false;
                #endregion

                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }





   

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

        private bool GetProperties(CExcel exls, Dictionary<string, string> cellMapping, ref Dictionary<string, object> buildingDefaultValues)
        {
            foreach (KeyValuePair<string, string> property in cellMapping)
            {
                try
                {
                    buildingDefaultValues.Add(property.Key, exls.GetCellValue(sheet, property.Value));
                }
                catch (Exception ex)
                {
                    SendErrorMessage(message: String.Format(ex.Message + "\t key = {0}", property.Key), sourceFunction: "GetProperties", exception: ex);
                    return false;
                }
            }
            return true;
        }

        private bool GetProperties(CExcel exls, KeyValuePair<string, string> cellMapping,
            ref Dictionary<string, object> buildingDefaultValues)
        {
            try
            {
                buildingDefaultValues.Add(cellMapping.Key,exls.GetCellValue(sheet,cellMapping.Value));
            }
            catch (Exception ex)
            {
                SendErrorMessage(message: String.Format(ex.Message + "\t key = {0}", cellMapping.Key), sourceFunction: "GetProperties", exception: ex);
                return false;
            }
            return true;
        }

        //Old system
        //void SetIspec(ref InputSpecification iSpec, Dictionary<string, string> propertyCellMapping)
        //{
        //    foreach (KeyValuePair<string, string> property in propertyCellMapping)
        //    {
        //        if (!iSpec.ContainsKey(property.Key))
        //            iSpec.Add(property.Key, new Number(property.Key));
        //    }
        //}
        
        //void SetInp(ref NonAtomic input, Dictionary<string, string> propertyCellMapping)
        //{
        //    foreach (KeyValuePair<string, string> property in propertyCellMapping)
        //    {
        //        if (!input.ContainsKey(property.Key))
        //            input.Add(property.Key, new Number(property.Key));
        //    }
        //}

        //void SetInp(ref NonAtomic input, KeyValuePair<string, string> property)
        //{
        //        if (!input.ContainsKey(property.Key))
        //            input.Add(property.Key, new Number(property.Key));
           
        //}
   
        private void Set(string sheet, string cell, object value, ref CExcel exls)
        {
            if (!exls.SetCellValue(sheet, cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {2} in sheet {3}", cell, value, sheet));
        }

        //bool GetSetBool(ref Dictionary<string, object> properties, string property)
        //{
        //    if (!properties.ContainsKey(property))
        //        properties.Add(property, false);

        //    if (properties[property] is bool)
        //        return (bool)properties[property];

        //    return false;
        //}

        //private bool Set(Feature building, Dictionary<string, string> propertyCellMappings, ref CExcel exls)
        //{
        //    foreach (KeyValuePair<string, string> property in propertyCellMappings)
        //    {
        //        if(!Set(building, property, ref exls))
        //            return false;
        //    }

        //    return true;
        //}

        //private bool Set(Feature building, KeyValuePair<string, string> propertyCellMapping, ref CExcel exls)
        //{
        //    //if (!CheckAndReportBuildingProp(building, propertyCellMapping.Key))
        //    //        return false;

        //    Set(sheet, propertyCellMapping.Value, building.properties[propertyCellMapping.Key], ref exls);                      

        //    return true;
        //}

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


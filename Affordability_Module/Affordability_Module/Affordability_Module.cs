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

        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_mdi,                 "C22"}
        };

        Dictionary<string, string> generalBuildingCellMapping = new Dictionary<string, string>()
        {
            {"size_of_household",                                           "C5"},
            {"size_of_dwelling",                                            "C6"},
            {"rent_or_cost_of_housing_per_month",                           "C10"},
            {"maximum_housing_costs_as_a_percentage_of_disposable_income",  "C17"}
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
            this.useBothVariantAndAsIS = false;
            this.newDashboardSystem = false;
            
            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_mdi };

            //Error handler
            this.ErrorRaised += CExcelModule_ErrorRaised;

            //Notification
            this.StatusMessage += CExcelModule_StatusMessage;            
        }
        
        protected override bool CalculateKpi(ModuleProcess process, CExcel exls, out Ecodistrict.Messaging.Data.Output output)
        {
            output = null;
            if (!KpiList.Contains(process.KpiId))
                return false;

            //if (process.KpiId == kpi_green)
            //    if (!SetProperties(process, exls, propertyCellMapping_Green))
            //        return false;

            //if ((process.KpiId == kpi_green) |
            //    (process.KpiId == kpi_biodiversity) |
            //    (process.KpiId == kpi_social_value) |
            //    (process.KpiId == kpi_climate_adaptation))
            //    if (!SetProperties(process, exls, propertyCellMapping_BSK))
            //        return false;

            //if ((process.KpiId == kpi_green) |
            //    (process.KpiId == kpi_social_value) |
            //    (process.KpiId == kpi_climate_adaptation))
            //    if (!SetProperties(process, exls, propertyCellMapping_SK))
            //        return false;

            //if ((process.KpiId == kpi_green) |
            //    (process.KpiId == kpi_biodiversity))
            //    if (!SetProperties(process, exls, propertyCellMapping_B))
            //        return false;

            //if ((process.KpiId == kpi_green) |
            //    (process.KpiId == kpi_social_value))
            //    if (!SetProperties(process, exls, propertyCellMapping_S))
            //        return false;

            //if ((process.KpiId == kpi_green) |
            //    (process.KpiId == kpi_climate_adaptation))
            //    if (!SetProperties(process, exls, propertyCellMapping_K))
            //        return false;

            output = new Ecodistrict.Messaging.Data.Output(process.KpiId);
            //output.KpiValue = exls.GetCellValue(sheet, kpiCellMapping[process.KpiId]) as double?;

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

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            var iSpec = new InputSpecification();

            if (!KpiList.Contains(kpiId))
                return null;

            if (kpiId == kpi_mdi)
            {
                NonAtomic geoJson = new GeoJson("buildings");
                SetInp(ref geoJson, generalBuildingCellMapping);
                geoJson.Add(include_household_electricity.Key, cb_include_household_electricity);
                geoJson.Add(include_domestic_hot_water.Key, cb_include_domestic_hot_water);
                geoJson.Add(include_cooling.Key , cb_include_cooling);
                geoJson.Add(include_space_heating.Key, cb_include_space_heating);
                geoJson.Add(useTypicalHouseholdElectricity.Key, householdElectricitySelect);
                geoJson.Add(manualHouseholdElectricity.Key, new Number(label: "If no: Enter annual consumption for household electricity", unit: "kWh"));
                geoJson.Add(energy_price_houshold_electricity.Key, new Number(label: "Enter the energy price for household electricity", unit: "EUR / kWh"));
                geoJson.Add(useTypicalDHW.Key, dHWSelect);
                geoJson.Add(manualDHW.Key, new Number(label: "If no: Enter annual consumption for domestic hot water", unit: "m\u00b3"));
                geoJson.Add(energy_price_domestic_hot_water.Key, new Number(label: "Enter the energy price for domestic hot water", unit: "EUR / m\u00b3"));
                geoJson.Add(manualCooling.Key, new Number(label: "Enter annual consumption for cooling", unit: "kWh"));
                geoJson.Add(energy_price_cooling.Key, new Number(label: "Enter the energy price for cooling", unit: "EUR / kWh"));
                geoJson.Add(manualSpaceHeating.Key, new Number(label: "Enter annual consumption for space heating", unit: "kWh"));
                geoJson.Add(energy_price_space_heating.Key, new Number(label: "Enter the energy price for space heating", unit: "EUR / kWh"));


                iSpec.Add("buildings", geoJson);
            }

            return iSpec;
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
            if (!CheckAndReportBuildingProp(building, propertyCellMapping.Key))
                    return false;

            Set(sheet, propertyCellMapping.Value, building.properties[propertyCellMapping.Key], ref exls);                      

            return true;
        }

        protected override bool CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls, out Ecodistrict.Messaging.Output.Outputs outputs)
        {
            outputs = null;
            double kpi = 0;

            if (!indata.ContainsKey("buildings"))
                return false;

            GeoJson buildingProperties = indata["buildings"] as GeoJson;

            if (buildingProperties == null)
                return false;

            int nrBuildings = 0;
            foreach (Feature building in buildingProperties.value.features)
            {
                Set(building, generalBuildingCellMapping, ref exls);

                #region Household Electricity
                if (building.properties.ContainsKey(include_household_electricity.Key))
                {
                    if (!Set(building, include_household_electricity, ref exls))
                        return false;

                    if ((bool)building.properties[include_household_electricity.Key])
                    {
                        if (!Set(building, energy_price_houshold_electricity, ref exls))
                            return false;

	                    if (!Set(building, useTypicalHouseholdElectricity, ref exls))
	                        return false;

                        if ((string)building.properties[useTypicalHouseholdElectricity.Key] == householdElectricityOpt.First().value)
                        {
	                        if (!Set(building, manualHouseholdElectricity, ref exls))
		                        return false;
                        }
                    }
                }
                #endregion

                #region DHW
                if (building.properties.ContainsKey(include_domestic_hot_water.Key))
                {
                    if (!Set(building, energy_price_domestic_hot_water, ref exls))
                        return false;

                    if (!Set(building, include_domestic_hot_water, ref exls))
                        return false;

                    if ((bool)building.properties[include_domestic_hot_water.Key])
                    {
                        if (!Set(building, useTypicalDHW, ref exls))
                            return false;

                        if ((string)building.properties[useTypicalDHW.Key] == dHWOpt.First().value)
                        {
                            if (!Set(building, manualDHW, ref exls))
                                return false;
                        }
                    }
                }
                #endregion


                #region Cooling
                if (building.properties.ContainsKey(include_cooling.Key))
                {
                    if (!Set(building, energy_price_cooling, ref exls))
                        return false;

                    if (!Set(building, include_cooling, ref exls))
                        return false;

                    if ((bool)building.properties[include_cooling.Key])
                    {
                       if (!Set(building, manualCooling, ref exls))
                            return false;
                    }
                }
                #endregion

                #region Space Heating
                if (building.properties.ContainsKey(include_space_heating.Key))
                {
                    if (!Set(building, energy_price_space_heating, ref exls))
                        return false;

                    if (!Set(building, include_space_heating, ref exls))
                        return false;

                    if ((bool)building.properties[include_space_heating.Key])
                    {
                        if (!Set(building, manualSpaceHeating, ref exls))
                            return false;
                    }
                }
                #endregion

                #region Result

                double resi = 100*Convert.ToDouble(exls.GetCellValue(sheetOutput, kpiCellMapping[kpiId]));
                kpi += resi;

                if (!building.properties.ContainsKey("kpiValue"))
                    building.properties.Add("kpiValue", resi);
                else
                    building.properties["kpiValue"] = Math.Round(resi, 0);

                #endregion

                ++nrBuildings;
            }
            kpi = kpi / (double)nrBuildings;


            outputs = new Ecodistrict.Messaging.Output.Outputs();

            switch (kpiId)
            {
                case kpi_mdi:
                    outputs.Add(new Ecodistrict.Messaging.Output.Kpi(Math.Round(kpi, 0), "Minimum disposable income", "EUR / month"));
                    break;
                default:
                    throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
            }

            Ecodistrict.Messaging.Output.GeoJson buildingsProps = new Ecodistrict.Messaging.Output.GeoJson(buildingProperties);
            outputs.Add(buildingsProps);

            return true;
        }

    }
}


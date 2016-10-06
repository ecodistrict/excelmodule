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

        private const string inputDistrictName = "District input for LCC";
        private const string inputBuildingName = "Building input for LCC";


        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_lcc,                 "H4"},
            {kpi_yearsToPayback,      "J4"},
            {kpi_totalLCC,      "H4"}
        };

        Dictionary<string, string> generalCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"discountrateexclinflation",                           "C4"},
            {"electricitypriceincrease",                            "C5"},
            {"districtheatingpriceincrease",                        "C6"},
            {"energypriceincreasedistrictheatingnaturalgas",        "C7"},
            {"energypriceincreaseotherfuel",                        "C8"},
            {"feedintariffpriceincrease",                           "C9"},
            {"lcacalcperiod",                                      "C10"},
            {"rentincrease",                                        "C11"},
            {"feedintariffprice",                                   "C12"},
            {"electricityprice",                                    "C13"},
            {"naturalgasprice",                                     "C14"},
            {"districtheatingprice",                                "C15"},
            {"otherfuelprice",                                      "C16"}
        };

        Dictionary<string, string> buildingCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"rentincomes",                         "C17"},
            {"fixedcostforelectricity",             "C24"},
            {"electricityconsumption",              "E25"},
            {"electricityproduction",               "E26"},
            {"fixedcostdistrictheating",            "C27"},       
            {"districtheatingconsumption",          "E28"},
            {"fixedcostfornaturalgas",              "C29"},
            {"naturalgasconsumption",               "E30"},
            {"fixedcostforotherfuel",               "C31"},
            {"otherfuelconsumption",                "E32"},
            {"operatingandmaintenancecosts",        "C34"}
        };

        Dictionary<string, string> heatingSystemCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"buildingheatingsystemlifeofproduct",                 "J24"}, 
            {"buildingheatingsystemheatsourceinitialinvestement",  "N24"},
            {"buildingheatingsystemheatsourceinstallationcost",    "O24"}
        };

        Dictionary<string, string> heatingPumpCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"buildingsystemheatpumplifeofproduct",      "J25"}, 
            {"buildingsystemheatpumpinitialinvestment",  "N25"},
            {"buildingsystemheatpumpinstallationcost",   "O25"}
        };

        Dictionary<string, string> heatingSystemBoreHoleCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"buildingsystemboreholelifeofproduct",          "J26"}, 
            {"buildingsystemboreholeinitialinvestment",      "N26"},
            {"buildingsystemboreholeinstallationcost",       "O26"}
        };

        Dictionary<string, string> heatingSystemCirculationPumpCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"heatingcirculationpumpsystemlifetime",               "J27"}, 
            {"heatingcirculationpumpsysteminvestmentcost",         "N27"},
            {"heatingcirculationpumpsysteminstallationcost",       "O27"}
        };

        Dictionary<string, string> buildingShellInsulationMaterial1CellMapping = new Dictionary<string, string>()
        {
            {"buildingshellinsulationmaterial1lifeofproduct",          "J28"}, 
            {"buildingshellinsulationmaterial1initialinvestement",     "N28"},
            {"buildingshellinsulationmaterial1installationcost",       "O28"}
        };

        Dictionary<string, string> buildingShellInsulationMaterial2CellMapping = new Dictionary<string, string>()
        {
            {"buildingshellinsulationmaterial2lifeofproduct",          "J29"}, 
            {"buildingshellinsulationmaterial2initialinvestement",     "N29"},
            {"buildingshellinsulationmaterial2installationcost",       "O29"}
        };

        Dictionary<string, string> buildingShellFacadeSystemCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"buildingshellfacadesystemlifetime",               "J30"}, 
            {"buildingshellfacadesysteminvestmentcost",         "N30"},
            {"buildingshellfacadesysteminstallationcost",       "O30"}
        };

        Dictionary<string, string> buildingShellWindowsCellMapping = new Dictionary<string, string>()
        {
            {"buildingshellwindowslifeofproduct",        "J31"}, 
            {"buildingshellwindowsinitialinvestement",   "N31"},
            {"buildingshellwindowsinstallationcost",     "O31"}
        };

        Dictionary<string, string> buildingShellDoorsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"buildingshelldoorssystemlifetime",               "J32"}, 
            {"buildingshelldoorssysteminvestmentcost",         "N32"},
            {"buildingshelldoorsysteminstallationcost",       "O32"}
        };

        Dictionary<string, string> VentilationSystemVentilationDuctsCellMapping = new Dictionary<string, string>()
        {
            {"ventilationsystemlifeofventilationducts",      "J33"}, 
            {"ventilationsysteminitialinvestementducts",     "N33"},
            {"ventilationsysteminstallationcostducts",       "O33"}
        };

        Dictionary<string, string> VentilationSystemAirflowAssemblyCellMapping = new Dictionary<string, string>()
        {
            {"ventilationsystemlifeofairflowassembly",                 "J34"}, 
            {"ventilationsysteminitialinvestementairflowassembly",     "N34"},
            {"ventilationsysteminstallationcostairflowassembly",       "O34"}
        };

        Dictionary<string, string> VentilationSystemDistributionHousingsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"ventilationsystemdistributionhousingslifetime",               "J35"}, 
            {"ventilationsystemdistributionhousingsinvestmentcost",         "N35"},
            {"ventilationsystemdistributionhousingsinstallationcost",       "O35"}
        };

        Dictionary<string, string> RadiatorsCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"buildingradiatorslifeofproduct",          "J36"}, 
            {"buildingradiatorsinitialinvestement",     "N36"},
            {"buildingradiatorsinstallationcost",       "O36"}
        };

        Dictionary<string, string> WaterTapsCellMapping = new Dictionary<string, string>()
        {
            {"watertapslifeofproduct",               "J37"}, 
            {"watertapsinitialinvestement",          "N37"},
            {"watertapsinstallationcost",            "O37"}
        };

        Dictionary<string, string> PipingSystemsCopperCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"pipingsystemscopperlifeofproduct",       "J38"}, 
            {"pipingsystemscopperinitialinvestement",  "N38"},
            {"pipingsystemscopperinstallationcost",    "O38"}
        };

        Dictionary<string, string> PipingSystemsPEXCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"pipingsystemspexlifetime",               "J39"}, 
            {"pipingsystemspexinvestmentcost",         "N39"},
            {"pipingsystemspexinstallationcost",       "O39"}
        };

        Dictionary<string, string> PipingSystemsPPCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"pipingsystemspplifetime",               "J40"}, 
            {"pipingsystemsppinvestmentcost",         "N40"},
            {"pipingsystemsppinstallationcost",       "O40"}
        };

        Dictionary<string, string> PipingSystemsCastIronCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"pipingsystemscastironlifetime",               "J41"}, 
            {"pipingsystemscastironinvestmentcost",         "N41"},
            {"pipingsystemscastironinstallationcost",       "O41"}
        };

        Dictionary<string, string> PipingSystemsGalvanisedSteelCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"pipingsystemsgalvanisedsteellifetime",               "J42"}, 
            {"pipingsystemsgalvanisedsteelinvestmentcost",         "N42"},
            {"pipingsystemsgalvanisedsteelinstallationcost",       "O42"}
        };

        Dictionary<string, string> PipingSystemsReliningCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"pipingsystemsrelininglifetime",               "J43"}, 
            {"pipingsystemsrelininginvestmentcost",         "N43"},
            {"pipingsystemsrelininginstallationcost",       "O43"}
        };

        Dictionary<string, string> ElectricalWiringCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"electricalwiringlifetime",               "J44"}, 
            {"electricalwiringinvestmentcost",         "N44"},
            {"electricalwiringinstallationcost",       "O44"}
        };


        Dictionary<string, string> EnergyProductionCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"energyproductionlifeofproduct",          "J45"}, 
            {"energyproductioninitialinvestement",     "N45"},
            {"energyproductioninstallationcost",       "O45"}
        };

        Dictionary<string, string> BuildingCondBoilersCellMapping = new Dictionary<string, string>()  //TODO update variable names
        {
            {"buildingcondboilerlifeofproduct",          "J46"}, 
            {"buildingcondboilerinitialinvestement",     "N46"},
            {"buildingcondboilerinstallationcost",       "O46"}
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

        private bool SetDistrictProperties(Dictionary<string, Object> currentData, CExcel exls, Dictionary<string, string> propertyCellMapping)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                //Dictionary<string, object> CurrentData = process.CurrentData;
                try
                {
                    {

                        if (currentData.ContainsKey(property.Key))
                        {
                            object value = currentData[property.Key];

                            double val = Convert.ToDouble(value);
                            if (val < 0)
                            {
                                //process.CalcMessage = String.Format("Property '{0}' has invalid data, only values equal or above zero is allowed; value: {1}", property.Key, val);
                                return false;
                            }

                            Set(sheet, property.Value, value, ref exls);
                        }
                        else
                        {
                            //process.CalcMessage = "";
                            return false;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    SendErrorMessage(
                        message:
                        String.Format(ex.Message + "\t key = {0}, isCurrentDataMissing = {1}", property.Key,
                            currentData == null), sourceFunction: "SetProperties", exception: ex);
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
                if (!CheckAndReportDistrictProp(process, process.CurrentData, inputDistrictName))
                    return false;


                if (!CheckAndReportDistrictProp(process,process.CurrentData, inputBuildingName))
                    return false;

                var nw = process.CurrentData[inputDistrictName] as List<Object>;
                var districtData = nw[0] as Dictionary<string, Object>;
                var myBuildings = process.CurrentData[inputBuildingName] as List<Object>;

                //Set common properties
                if (!SetDistrictProperties(districtData, exls, generalCellMapping))
                    return false;

                //Calculate kpi
                //NEW_CODE: Start with getting all default (Inital start) values from the excel Sheet for buildingdata
                Dictionary<string, object> buildingDefaultValues = new Dictionary<string, object>();
                if (myBuildings != null && myBuildings.Count > 0)
                    if (!GetBuildingDefaultValues(exls, out buildingDefaultValues))
                        return false;

                outputDetailed = new Ecodistrict.Messaging.Data.OutputDetailed(process.KpiId);
                double kpiValue = 0;
                int noRenovatedBuildings = 0;
                
                foreach (Dictionary<string,object> buildingData in myBuildings)
                {
                    double kpiValuei;
                    bool changesMade;
                    if (!SetInputDataOneBuilding(buildingData, exls, out changesMade))
                        return false;
                    kpiValuei = Convert.ToDouble(exls.GetCellValue(sheet, kpiCellMapping[process.KpiId]));
                    if (changesMade)
                        ++noRenovatedBuildings;

                    if(noRenovatedBuildings % 50==0)
                        SendStatusMessage(string.Format("{0} building processed", noRenovatedBuildings));

                    //NEW_CODE: Reset all used building values
                    if (buildingDefaultValues != null &&
                        !SetInputDataOneBuilding(buildingDefaultValues, exls, out changesMade))
                        return false;

                    kpiValue += kpiValuei;
                    //TODO fix this below
                    outputDetailed.KpiValueList.Add(new Ecodistrict.Messaging.Data.GeoObject("building", buildingData["building_id"] as string, process.KpiId, kpiValuei));

                }


                if (noRenovatedBuildings > 0 && (process.KpiId != kpi_totalLCC))
                    kpiValue /= Convert.ToDouble(noRenovatedBuildings);

                output = new Ecodistrict.Messaging.Data.Output(process.KpiId, Math.Round(kpiValue, 1));

                SendStatusMessage(string.Format("Totally {0} building processed", noRenovatedBuildings));

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
                if (!GetProperties(exls, heatingSystemCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, heatingPumpCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, heatingSystemBoreHoleCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, heatingSystemCirculationPumpCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, buildingShellInsulationMaterial1CellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, buildingShellInsulationMaterial2CellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, buildingShellFacadeSystemCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, buildingShellWindowsCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, buildingShellDoorsCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, VentilationSystemVentilationDuctsCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, VentilationSystemAirflowAssemblyCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, VentilationSystemDistributionHousingsCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, RadiatorsCellMapping, ref buildingDefaultValues))
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
                if (!GetProperties(exls, EnergyProductionCellMapping, ref buildingDefaultValues))
                    return false;
                if (!GetProperties(exls, BuildingCondBoilersCellMapping, ref buildingDefaultValues))
                    return false;
                #endregion

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private bool SetInputDataOneBuilding(Dictionary<string, object> buildingData, CExcel exls, out bool changesMade)
        {
            changesMade = false;
            bool changesMade_i = false;

            try
            {
                #region Set Data
                if (!SetProperties(buildingData, exls, buildingCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, heatingSystemCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, heatingPumpCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, heatingSystemBoreHoleCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, heatingSystemCirculationPumpCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, buildingShellInsulationMaterial1CellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, buildingShellInsulationMaterial2CellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, buildingShellFacadeSystemCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, buildingShellWindowsCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, buildingShellDoorsCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, VentilationSystemVentilationDuctsCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, VentilationSystemAirflowAssemblyCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, VentilationSystemDistributionHousingsCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, RadiatorsCellMapping, out changesMade_i))
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

                if (!SetProperties(buildingData, exls, EnergyProductionCellMapping, out changesMade_i))
                    return false;
                changesMade = changesMade | changesMade_i;

                if (!SetProperties(buildingData, exls, BuildingCondBoilersCellMapping, out changesMade_i))
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


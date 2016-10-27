using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;
using Ecodistrict.Messaging.Data;

namespace RenobuildModule.SourceCode
{
    class RenobuildModuleNew: CExcelModule
    {
        // - Kpis
        const string kpi_gwp = "change-of-global-warming-potential";
        const string kpi_gwp_per_heated_area = "change-of-global-warming-potential-per-heated-area";
        const string kpi_peu = "change-of-primary-energy-use";
        const string kpi_peu_per_heated_area = "change-of-primary-energy-use-per-heated-area";
        const string sheet = "Indata";
        const string buidingIdKey = "gml_id";

        private const string inputDistrictName = "District input for LCA";
        private const string inputBuildingName = "Building input for LCA";


        
        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>() //TMP
        {
            {kpi_gwp,                   "C31"},
            {kpi_gwp_per_heated_area,   "C31"},
            {kpi_peu,                   "C32"},
            {kpi_peu_per_heated_area,   "C32"}
        };

        private Dictionary<string, string> districtCellMapping = new Dictionary<string, string>()
        {
            {"LCACalcPeriod",            "C16"},            {"LcaCountry",               "C17"},            {"LcaGwpDistrictHeating",    "C20"},            {"LcaPefDistrictHeating",    "C21"}
        };

        private Dictionary<string, string> buildingCellMapping = new Dictionary<string, string>()
        {
            {"HeatedFloorArea",                                     "C25"},            {"LcaHeatSourceBeforeRenovation",                       "C93"},            {"LcaHeatDemandAfterRenovation",                        "C95"},            {"LcaElectricityDemandHeatPumpAfterRenovation",         "C96"},            {"LcaHeatSourceAfterRenovation",                        "C94"},            {"BuildingHeatingSystemLifeOfProduct",                  "C100"},            {"LcaHeatSourceDesignCapacity",                         "C103"},            {"LcaHeatSourceWeight",                                 "C104"},            {"LcaGeothermalBoreholeDepth",                          "C109"},            {"HeatingCirculationPumpSystemLifeTime",                "C114"},            {"LcaPumpWeight",                                       "C118"},            {"LcaChangeHeatDemand",                                 "C298"},            {"LcaChangeElectricityDemandExclHeatCool",              "C299"},            {"LcaChangeElectricityDemandCooling",                   "C300"},            {"BuildingShellInsulationMaterial1LifeOfProduct",       "C127"},            {"LcaInsulationMaterial1",                              "C128"},            {"LcaInsulationMaterial1Amount",                        "C130"},            {"BuildingShellInsulationMaterial2LifeOfProduct",       "C138"},            {"LcaInsulationMaterial2",                              "C139"},            {"LcaInsulationMaterial2Amount",                        "C141"},            {"BuildingShellFacadeSystemLifeTime",                   "C149"},            {"LcaFacadeSystem",                                     "C150"},            {"LcaFacadeSystemArea",                                 "C152"},            {"BuildingShellWindowsLifeOfProduct",                   "C160"},            {"LcaWindowType",                                       "C161"},            {"LcaWindowArea",                                       "C163"},            {"BuildingShellDoorsSystemLifeTime",                    "C171"},            {"LcaDoorType",                                         "C172"},            {"LcaDoorsNumber",                                      "C174"},            {"VentilationSystemLifeOfVentilationDucts",             "C183"},            {"LcaVentilationDuctsMaterial",                         "C184"},            {"LcaVentilationDuctsWeight",                           "C185"},            {"VentilationSystemLifeOfVentilationUnit",              "C193"},            {"LcaVenitlationUnitType",                              "C194"},            {"LcaVentilationUnitDesignExhaustAirFlow",              "C195"},            {"VentilationSystemDistributionHousingsLifeTime",       "C203"},            {"LcaAirDistributionHousingsNumber",                    "C208"},            {"BuildingRadiatorsLifeOfProduct",                      "C216"},            {"LcaRadiatorsType",                                    "C217"},            {"LcaRadiatorsWeight",                                  "C218"},            {"WaterTapsLifeOfProduct",                              "C226"},            {"LcaWaterTapsNumber",                                  "C228"},            {"PipingSystemsCopperLifeOfProduct",                    "C236"},            {"LcaCopperPipesWeight",                                "C237"},            {"PipingSystemsPEXLifeTime",                            "C245"},            {"LcaPEXPipesWeight",                                   "C246"},            {"PipingSystemsPPLifeTime",                             "C254"},            {"LcaPPPipesWeight",                                    "C255"},            {"PipingSystemsCastIronLifeTime",                       "C263"},            {"LcaCastIronPipesWeight",                              "C264"},            {"PipingSystemsGalvanisedSteelLifeTime",                "C272"},            {"LcaGalvSteelPipesWeight",                             "C273"},            {"PipingSystemsReliningLifeTime",                       "C281"},            {"LcaReliningMaterialWeight",                           "C282"},            {"ElectricalWiringLifeTime",                            "C290"},            {"LcaElectricalWiresWeight",                            "C291"},            {"EnergyProductionLifeOfProduct",                       "C304"},            {"LcaEnergyProductionFacility",                         "C305"},            {"LcaEnergyProductionSize",                             "C306"},            {"LcaElectricityProduction",                            "C307"}
        };

        #endregion

        public RenobuildModuleNew()
        {
            useDummyDB = false;
            this.useBothVariantAndAsISForVariant = false; //??

            this.KpiList = kpiCellMapping.Keys.ToList();

            //Error handler
            this.ErrorRaised += CExcelModule_ErrorRaised;

            //Notifycation
            this.StatusMessage += CExcelModule_StatusMessage;
        }

        private bool SetProperties(Dictionary<string, object> buildingData, CExcel exls, Dictionary<string, string> propertyCellMapping, out bool changesMade)
        {

            #if(ToClipBoard)
            using (FileStream fs = File.Open(@"C:\Temp\EcoTemp\BuildingData.csv", FileMode.Append))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    foreach (KeyValuePair<string, string> pair in propertyCellMapping)
                    {
                        sw.WriteLineAsync(string.Format("{0}\t{1}\t{2}", pair.Key, pair.Value, buildingData[pair.Key]));
                    }
                }
            }
            #endif


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
            #if(ToClipBoard)
            using (FileStream fs = File.Open(@"C:\Temp\EcoTemp\DistData.csv", FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    foreach (KeyValuePair<string, object> pair in currentData)
                    {
                        sw.WriteLineAsync(string.Format("{0}\t{1}\t{2}", pair.Key, propertyCellMapping[pair.Key],pair.Value));
                    }
                }
            }
            #endif

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

        private bool GetBuildingDefaultValues(CExcel exls, out Dictionary<string, object> buildingDefaultValues)
        {
            buildingDefaultValues = new Dictionary<string, object>();
            try
            {

                if (!GetProperties(exls, buildingCellMapping, ref buildingDefaultValues))
                    return false;

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



        protected override bool CalculateKpi(ModuleProcess process, CExcel exls, out Ecodistrict.Messaging.Data.Output output, out Ecodistrict.Messaging.Data.OutputDetailed outputDetailed)
        {
            output = null;
            outputDetailed = null;
            bool perHeatedArea=false;

            if (!KpiList.Contains(process.KpiId))
            {
                process.CalcMessage = String.Format("kpi not available for this module, requested kpi: {0}", process.KpiId);
                return false;
            }

            switch (process.KpiId)
            {
                case kpi_gwp:
                case kpi_peu:
                    break;
                case kpi_gwp_per_heated_area:
                case kpi_peu_per_heated_area:
                    perHeatedArea = true;
                    break;
            }

            if(!CheckAndReportDistrictProp(process,process.CurrentData,inputDistrictName))
                return false;

            if(!CheckAndReportBuildingProp(process,process.CurrentData,inputBuildingName))
                return false;

            var nw = process.CurrentData[inputDistrictName] as List<object>;
            var districtdata = nw[0] as Dictionary<string, object>;

            var myBuildings = process.CurrentData[inputBuildingName] as List<object>;

            //Set common properties
            if(!SetDistrictProperties(districtdata,exls,districtCellMapping))
                return false;

            //Get all default building data data from Excel document
            var buildingDefaultValues=new Dictionary<string, object>();
            if(myBuildings!=null && myBuildings.Count>0)
                if (!GetBuildingDefaultValues(exls, out buildingDefaultValues))
                    return false;

            outputDetailed=new OutputDetailed(process.KpiId);
            double kpiValue = 0;
            int noOfRenovatedBuildings=0;

            foreach (Dictionary<string,object> buildingData in myBuildings)
            {
                double kpiValuei;
                bool changesMade;

                if(!SetInputDataOneBuilding(buildingData,exls,out changesMade))
                    return false;

                kpiValuei = Convert.ToDouble(exls.GetCellValue(sheet, kpiCellMapping[process.KpiId]));
                if (changesMade)
                    noOfRenovatedBuildings++;

                if(noOfRenovatedBuildings%50==0)
                    SendStatusMessage(string.Format("{0} building processed", noOfRenovatedBuildings));

                //Reset buildingdata do tefault
                if(buildingDefaultValues!=null &&
                    !SetInputDataOneBuilding(buildingDefaultValues,exls,out changesMade))
                    return false;

                if (perHeatedArea)
                    kpiValuei *= 1000; //From tonnes CO2 eq / m2 to kg CO2 eq/ m2 and MWh / m2 to kWh/ m2 resp.

                kpiValue += kpiValuei;

                outputDetailed.KpiValueList.Add(new GeoObject("building",buildingData["building_id"] as string,process.KpiId,kpiValuei));
            }

            if (noOfRenovatedBuildings > 0 & (process.KpiId == kpi_gwp | process.KpiId == kpi_peu))
                kpiValue /= 30.0 * Convert.ToDouble(noOfRenovatedBuildings);
            else if (process.KpiId == kpi_gwp_per_heated_area | process.KpiId == kpi_peu_per_heated_area)
                kpiValue /= 30.0 * Convert.ToDouble(250000); //TMP

            output = new Ecodistrict.Messaging.Data.Output(process.KpiId, Math.Round(kpiValue, 1));

            return true;



        }
    }
}

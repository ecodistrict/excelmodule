using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace MobilityModule
{
    class MobilityModule : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_modalsplit_origin_private_transport        = "modal-split-origin-private-transport";
        const string kpi_modalsplit_origin_public_transport         = "modal-split-origin-public-transport";
        const string kpi_modalsplit_origin_slow_modes               = "modal-split-origin-slow-modes";
        const string kpi_modalsplit_destination_private_transport   = "modal-split-destination-private-transport";
        const string kpi_modalsplit_destination_public_transport    = "modal-split-destination-public-transport";
        const string kpi_modalsplit_destination_slow_modes          = "modal-split-destination-slow-modes";
        const string sheetInput = "Calculations";
        const string sheetOutput = "Measures";
        const string sheetSettings = "Measures";
        const string kpiDB_modalsplit_origin_private_transport      = "Modal split - origin - Private transport";
        const string kpiDB_modalsplit_origin_public_transport       = "Modal split - origin - Public transport";
        const string kpiDB_modalsplit_origin_slow_modes             = "Modal split - origin - Slow modes";
        const string kpiDB_modalsplit_destination_private_transport = "Modal split - destination - Private transport";
        const string kpiDB_modalsplit_destination_public_transport  = "Modal split - destination - Public transport";
        const string kpiDB_modalsplit_destination_slow_modes        = "Modal split - destination - Slow modes";

        private const string inputDistrictName = "District input for Mobility";


        #region Cell Mapping


        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_modalsplit_origin_private_transport,       "K24"},
            {kpi_modalsplit_origin_public_transport,        "L24"},
            {kpi_modalsplit_origin_slow_modes,              "M24"},
            {kpi_modalsplit_destination_private_transport,  "K25"},
            {kpi_modalsplit_destination_public_transport,   "L25"},
            {kpi_modalsplit_destination_slow_modes,         "M25"}
        };


        Dictionary<string, string> kpi_kpi_mapping= new Dictionary<string, string>()
        {
            {kpi_modalsplit_origin_private_transport,       kpiDB_modalsplit_origin_private_transport},
            {kpi_modalsplit_origin_public_transport,        kpiDB_modalsplit_origin_public_transport},
            {kpi_modalsplit_origin_slow_modes,              kpiDB_modalsplit_origin_slow_modes},
            {kpi_modalsplit_destination_private_transport,  kpiDB_modalsplit_destination_private_transport},
            {kpi_modalsplit_destination_public_transport,   kpiDB_modalsplit_destination_public_transport},
            {kpi_modalsplit_destination_slow_modes,         kpiDB_modalsplit_destination_slow_modes}
        };

        //Needed AsIs properties
        //Dictionary<string, string> propertyCellMapping_AsIs = new Dictionary<string, string>()
        //{
        //    {kpiDB_modalsplit_origin_private_transport,         "H11"},
        //    {kpiDB_modalsplit_origin_public_transport,          "I11"},
        //    {kpiDB_modalsplit_origin_slow_modes,                "J15"},
        //    {kpiDB_modalsplit_destination_private_transport,    "H12"},
        //    {kpiDB_modalsplit_destination_public_transport,     "I12"},
        //    {kpiDB_modalsplit_destination_slow_modes,           "J16"}
        //};

        private Dictionary<string, string> propertyCellMapping_CalcSheet = new Dictionary<string, string>()
        {
            {"M0001ModalSplitOriginPrivateTransport",                                   "H11"},
            {"M0002ModalSplitOriginPublicTransport",                                    "I11"},
            {"M0003ModalSplitOriginSlowModes",                                          "J15"},
            {"M0004ModalSplitDestinationPrivateTransport",                              "H12"},
            {"M0005ModalSplitDestinationPublicTransport",                               "I12"},
            {"M0006ModalSplitDestinationSlowModes",                                     "J15"},
            {"M0007TotalFloorAreaHousing",                                              "F20"},
            {"M0008TotalFloorAreaOffices",                                              "F21"},
            {"M0014OccupancyOffices",                                                   "H21"},
            {"M0009TotalFloorAreaRetail",                                               "F22"},
            {"M0015OccupancyRetail",                                                    "H22"},
            {"M0010TotalFloorAreaIndustrial",                                           "F23"},
            {"M0016OccupancyIndustrial",                                                "H23"},
            {"M0011TotalFloorAreaOther",                                                "F24"},
            {"M0017OccupancyOther",                                                     "H24"},
            {"M0012TotalFloorAreaParkingArea",                                          "F25"},
            {"M0018OccupancyParkingArea",                                               "H25"},
            {"M0101ImpactWithoutMeasure",                                               "F36"},
            {"M0102ImpactOnPrivateTransport",                                           "H36"},
            {"M0103ImpactOnPublicTransport",                                            "I36"},
            {"M0104ImpactOnSlowModes",                                                  "J36"},
            {"M0201ImpactWithoutMeasure",                                               "F42"},
            {"M0202ImpactOnPrivateTransport",                                           "H42"},
            {"M0013OccupancyHousing",                                                   "H20"},
            {"M0203ImpactOnPublicTransport",                                            "I42"},
            {"M0204ImpactOnSlowModes",                                                  "J42"},
            {"M0301ImpactWithoutMeasure",                                               "F49"},
            {"M0302ImpactOnPrivateTransport",                                           "H49"},
            {"M0303ImpactOnPublicTransport",                                            "I49"},
            {"M0304ImpactOnSlowModes",                                                  "J49"},
            {"M0401ImpactWithoutMeasure",                                               "F57"},
            {"M0402ImpactOnPrivateTransport",                                           "H57"},
            {"M0403ImpactOnPublicTransport",                                            "I57"},
            {"M0404ImpactOnSlowModes",                                                  "J57"},
            {"M0501ImpactWithoutMeasure",                                               "F64"},
            {"M0502ImpactOnPrivateTransport",                                           "H64"},
            {"M0503ImpactOnPublicTransport",                                            "I64"},
            {"M0504ImpactOnSlowModes",                                                  "J64"},
            {"M0601PercentageOfCarsParkedOutsiteParkingAreas",                          "H75"},
            {"M0602CarOccupancy",                                                       "H76"},
            {"M0603PercentageOfPeopleWhoWillShift",                                     "H78"},
            {"M0604ImpactWithoutMeasure",                                               "F83"},
            {"M0605ImpactOnSlowModes",                                                  "J83"},
            {"M0701ImpactWithoutMeasure",                                               "F93"},
            {"M0702ImpactOnPrivateTransport",                                           "H93"},
            {"M0703ImpactOnPublicTransport",                                            "I93"},
            {"M0704ImpactOnSlowModes",                                                  "J93"},
            {"M0705ImpactWithoutMeasure",                                               "F100"},
            {"M0706ImpactOnPrivateTransport",                                           "H100"},
            {"M0707ImpactOnPublicTransport",                                            "I100"},
            {"M0708ImpactOnSlowModes",                                                  "J100"},
            {"M0801DaysPerWeekThatEmployeesWorkFromHome",                               "H114"},
            {"M0802PercentFloorAreaOffBuildFlexwork",                                   "H115"},
            {"M0803ImpactWithoutMeasure",                                               "F120"},
            {"M0804ImpactOnSlowModes",                                                  "J120"},
            {"M0901ImpactWithoutMeasure",                                               "F128"},
            {"M0902ImpactOnPrivateTransport",                                           "H128"},
            {"M0903ImpactOnPublicTransport",                                            "I128"},
            {"M0904ImpactOnSlowModes",                                                  "J128"},
            {"M1001ImpactWithoutMeasure",                                               "F138"},
            {"M1002ImpactOnPrivateTransport",                                           "H138"},
            {"M1003ImpactOnPublicTransport",                                            "I138"},
            {"M1004ImpactOnSlowModes",                                                  "J138"}
        };

        private Dictionary<string, string> propertyCellMapping_MeasureSheet = new Dictionary<string, string>()
        {
            {"M01CombineTramAndBusInfrastructure",                                                    "E8"},
            {"M02LargerTramAndBusVehiclesUpToMax20PerCentIncrease",                                   "E9"},
            {"M03HiFreqTramBusServ20PerCentIncrStopServ",                                             "E10"},
            {"M04OptimisationOfBusRoute",                                                             "E11"},
            {"M05ModificationOfTramAndBusRoutesToConnectToPAndR",                                     "E12"},
            {"M06ParkingZonePolicy ",                                                                 "E13"},
            {"M07PAndR ",                                                                             "E14"},
            {"M08FlexWorking",                                                                        "E15"},
            {"M09PromotionOfPublicTransportEmployersPayingForPublicTransport",                        "E16"},
            {"M10MixedUsePlanning",                                                                   "E17"}
        };
        
        
#region Old Code
        

        //General data
        //Dictionary<string, string> propertyCellMapping_General = new Dictionary<string, string>()
        //{
        //    {"Total Floor area - Housing",        "F20"},
        //    {"Total Floor area - Offices",        "F21"},
        //    {"Total Floor area - Retail",         "F22"},
        //    {"Total Floor area - Industrial",     "F23"},
        //    {"Total Floor area - Other",          "F24"},
        //    {"Total Floor area - Parking area",   "F25"},
        //    {"Occupancy - Housing",               "H20"},
        //    {"Occupancy - Offices",               "H21"},
        //    {"Occupancy - Retail",                "H22"},
        //    {"Occupancy - Industrial",            "H23"},
        //    {"Occupancy - Other",                 "H24"},
        //    {"Occupancy - Parking area",          "H25"}
        //};

        ////1	Public transport	Combine tram and bus infrastructure
        //KeyValuePair<string, string> propertyCellMapping_Use01 =
        //    new KeyValuePair<string, string>("Combine tram and bus infrastructure", "E8");
        //Dictionary<string, string> propertyCellMapping_01= new Dictionary<string, string>()
        //{                                
        //    {"Combine tram and bus infrastructure - Impact without measure",       "F36"},
        //    {"Combine tram and bus infrastructure - Impact on private transport",  "H36"},
        //    {"Combine tram and bus infrastructure - Impact on public transport",   "I36"},
        //    {"Combine tram and bus infrastructure - Impact on slow modes",         "J36"}
        //};

        ////2	Public transport	Larger tram and bus vehicles up to max. 20% increase
        //KeyValuePair<string, string> propertyCellMapping_Use02 =
        //    new KeyValuePair<string, string>("Larger tram and bus vehicles", "E9");
        //Dictionary<string, string> propertyCellMapping_02 = new Dictionary<string, string>()
        //{                               
        //    {"Larger tram and bus vehicles - Impact without measure",       "F42"},
        //    {"Larger tram and bus vehicles - Impact on private transport",  "H42"},
        //    {"Larger tram and bus vehicles - Impact on public transport",   "I42"},
        //    {"Larger tram and bus vehicles - Impact on slow modes",         "J42"}
        //};

        ////3	Public transport	Higher frequency tram and bus services up to max. 20% increase of tram and bus stop service.
        //KeyValuePair<string, string> propertyCellMapping_Use03 =
        //    new KeyValuePair<string, string>("Higher frequency tram and bus services", "E10");
        //Dictionary<string, string> propertyCellMapping_03 = new Dictionary<string, string>()
        //{                                                          
        //    {"Higher frequency tram and bus services - Impact without measure",       "F49"},
        //    {"Higher frequency tram and bus services - Impact on private transport",  "H49"},
        //    {"Higher frequency tram and bus services - Impact on public transport",   "I49"},
        //    {"Higher frequency tram and bus services - Impact on slow modes",         "J49"}
        //};

        ////4	Public transport	Optimization of bus route
        //KeyValuePair<string, string> propertyCellMapping_Use04 =
        //    new KeyValuePair<string, string>("Optimisation of bus route", "E11");
        //Dictionary<string, string> propertyCellMapping_04 = new Dictionary<string, string>()
        //{                                                           
        //    {"Optimisation of bus route - Impact without measure",       "F57"},
        //    {"Optimisation of bus route - Impact on private transport",  "H57"},
        //    {"Optimisation of bus route - Impact on public transport",   "I57"},
        //    {"Optimisation of bus route - Impact on slow modes",         "J57"}
        //};

        ////5	Public transport	Modification of tram and bus routes to connect to P&R
        //KeyValuePair<string, string> propertyCellMapping_Use05 =
        //    new KeyValuePair<string, string>("Modification of tram and bus routes to connect to P and R", "E12");
        //Dictionary<string, string> propertyCellMapping_05 = new Dictionary<string, string>()
        //{                                                         
        //    {"Modification of tram and bus routes to connect to P and R - Impact without measure",       "F64"},
        //    {"Modification of tram and bus routes to connect to P and R - Impact on private transport",  "H64"},
        //    {"Modification of tram and bus routes to connect to P and R - Impact on public transport",   "I64"},
        //    {"Modification of tram and bus routes to connect to P and R - Impact on slow modes",         "J64"}
        //};

        ////6	Private transport	Parking zone policy 
        //KeyValuePair<string, string> propertyCellMapping_Use06 =
        //    new KeyValuePair<string, string>("Parking zone policy", "E13");
        //Dictionary<string, string> propertyCellMapping_06 = new Dictionary<string, string>()
        //{                                         
        //    {"Parking zone policy - Percentage of cars parked outsite parking areas",         "H75"},
        //    {"Parking zone policy - car occupancy ",                                           "H76"},      
        //    {"Parking zone policy - Percentage of people who will shift from car to public transport with strict parking policy", "H78"},                                                                  
        //    {"Parking zone policy - Impact without measure",       "F83"},
        //    {"Parking zone policy - Impact on slow modes",         "J83"} 
        //};

        ////7	Private transport	P&R
        //KeyValuePair<string, string> propertyCellMapping_Use07 =
        //    new KeyValuePair<string, string>("P and R", "E14");
        //Dictionary<string, string> propertyCellMapping_07_1 = new Dictionary<string, string>() //If 6 is used (Y)
        //{                                                           
        //    {"P and R - Impact without measure when using parking zone policy",       "F93"}, 
        //    {"P and R - Impact on private transport when using parking zone policy",  "H93"},
        //    {"P and R - Impact on public transport when using parking zone policy",   "I93"},
        //    {"P and R - Impact on slow modes when using parking zone policy",         "J93"}
        //};
        //Dictionary<string, string> propertyCellMapping_07_2 = new Dictionary<string, string>() //If 6 is not used (N)
        //{                                                           
        //    {"P and R - Impact without measure",       "F100"}, 
        //    {"P and R - Impact on private transport",  "H100"},
        //    {"P and R - Impact on public transport",   "I100"},
        //    {"P and R - Impact on slow modes",         "J100"}
        //};

        ////8	Traffic management	Flex working
        //KeyValuePair<string, string> propertyCellMapping_Use08 =
        //    new KeyValuePair<string, string>("Flex working", "E15");
        //Dictionary<string, string> propertyCellMapping_08 = new Dictionary<string, string>()
        //{                   
        //    {"Flex working - Days per week that employees work from home",                             "H114"},               
        //    {"Flex working - Percentage of floor area office buildings participating in flexworking",  "H115"},                                                        
        //    {"Flex working - Impact without measure",                                                  "F120"},
        //    {"Flex working - Impact on slow modes",                                                    "J120"} 
        //};

        ////9	Traffic management	Promotion of public transport (employers paying for public transport)
        //KeyValuePair<string, string> propertyCellMapping_Use09 =
        //    new KeyValuePair<string, string>("Promotion of public transport", "E16");
        //Dictionary<string, string> propertyCellMapping_09 = new Dictionary<string, string>()
        //{                                                           
        //    {"Promotion of public transport - Impact without measure",       "F128"},
        //    {"Promotion of public transport - Impact on private transport",  "H128"},
        //    {"Promotion of public transport - Impact on public transport",   "I128"},
        //    {"Promotion of public transport - Impact on slow modes",         "J128"}
        //};

        ////10	Traffic management	Mixed use planning
        //KeyValuePair<string, string> propertyCellMapping_Use10 =
        //    new KeyValuePair<string, string>("Mixed use planning", "E17");
        //Dictionary<string, string> propertyCellMapping_10 = new Dictionary<string, string>()
        //{                                                            
        //    {"Mixed use planning - Impact without measure",       "F138"},
        //    {"Mixed use planning - Impact on private transport",  "H138"},
        //    {"Mixed use planning - Impact on public transport",   "I138"},
        //    {"Mixed use planning - Impact on slow modes",         "J138"}
        //};

        #endregion

        #endregion

        #endregion

        public MobilityModule()
        {
            useDummyDB = false;
            useXLSData = false;
            useBothVariantAndAsISForVariant = true;

            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "";
            this.ModuleName = "MobilityModule";

            //List of kpis the module can calculate
            this.KpiList = kpiCellMapping.Keys.ToList();

            //Error handler
            this.ErrorRaised += CExcelModule_ErrorRaised;

            //Notification
            this.StatusMessage += CExcelModule_StatusMessage;
        }

        private void Set(string sheet, string cell, object value, ref CExcel exls)
        {
            if (!exls.SetCellValue(sheet, cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {2} in sheet {3}", cell, value, sheet));
        }
        
        protected override bool CalculateKpi(ModuleProcess process, CExcel exls, out Ecodistrict.Messaging.Data.Output output, out Ecodistrict.Messaging.Data.OutputDetailed outputDetailed)
        {
            try
            {
                output = null;
                outputDetailed = null;

                if (!KpiList.Contains(process.KpiId))
                {
                    process.CalcMessage = "Kpi not available for this module";
                    return false;
                }

                if (!CheckAndReportDistrictProp(process, process.CurrentData, inputDistrictName))
                    return false;

                string dbKpiId = kpi_kpi_mapping[process.KpiId];
                
                double kpiValue = 0;

                var nw = process.CurrentData[inputDistrictName] as List<Object>;
                var districtData = nw[0] as Dictionary<string, Object>;

                if (!SetDistrictProperties(districtData, exls, propertyCellMapping_CalcSheet, sheetInput))
                {
                    return false;
                }

                if (!SetDistrictProperties(districtData, exls, propertyCellMapping_MeasureSheet, sheetSettings))
                {
                    return false;
                }
                
                double? val = exls.GetCellValue(sheetOutput, kpiCellMapping[process.KpiId]) as double?;

                if (val == null)
                        return false;

                kpiValue = Math.Round((double)val,1);

                output = new Ecodistrict.Messaging.Data.Output(process.KpiId, kpiValue);

                return true;

                #region Old Code

                //Prepare AsIs data
                //if (process.As_IS_Data == null)
                //{
                //    process.CalcMessage = "No as is data";
                //    return false;
                //}

                ////string distrName = "District";
                //string distrName = "District trafic area";

                //if (!process.As_IS_Data.ContainsKey(distrName))
                //{
                //    process.CalcMessage = "As is district information missing";
                //    return false;
                //}

                //Dictionary<string, object> dataAsIS ;
                //if (process.As_IS_Data[distrName] is Dictionary<string, object>)
                //    dataAsIS = process.As_IS_Data[distrName] as Dictionary<string, object>;
                //else
                //{
                //    process.CalcMessage = "As is data received from data module is wrongly formated";
                //    return false;
                //}


                ////AsIS
                //if (process.IsAsIS)
                //{                   
                //    if (!dataAsIS.ContainsKey(dbKpiId))
                //    {
                //        process.CalcMessage = "As is information missing";
                //        return false;
                //    }

                //    kpiValue = Convert.ToDouble(dataAsIS[dbKpiId]);
                //}
                ////Variant
                //else
                //{
                //    //Prepare Variant data
                //    if (process.Variant_Data == null)
                //    {
                //        process.CalcMessage = "No variant data";
                //        return false;
                //    }

                //    if (!process.Variant_Data.ContainsKey(distrName))
                //    {
                //        process.CalcMessage = "Variant district information missing";
                //        return false;
                //    }


                //    Dictionary<string, object> dataVariant;
                //    if (process.Variant_Data[distrName] is Dictionary<string, object>)
                //        dataVariant = process.Variant_Data[distrName] as Dictionary<string, object>;
                //    else
                //    {
                //        process.CalcMessage = "Variant data received from data module is wrongly formated";
                //        return false;
                //    }
                    

                //    //Set Data
                //    if (!SetProperties(process, dataAsIS, exls, propertyCellMapping_AsIs))
                //        return false;

                //    if (!SetProperties(process, dataVariant, exls, propertyCellMapping_General))
                //        return false;

                //    //01
                //    bool used01;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use01, propertyCellMapping_01, exls, out used01))
                //        return false;

                //    //02
                //    bool used02;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use02, propertyCellMapping_02, exls, out used02))
                //        return false;

                //    //03
                //    bool used03;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use03, propertyCellMapping_03, exls, out used03))
                //        return false;

                //    //04
                //    bool used04;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use04, propertyCellMapping_04, exls, out used04))
                //        return false;

                //    //05
                //    bool used05;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use05, propertyCellMapping_05, exls, out used05))
                //        return false;

                //    //06
                //    bool used06;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use06, propertyCellMapping_06, exls, out used06))
                //        return false;

                //    //07
                //    bool used07;
                //    if (used06)
                //    {
                //        if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use07, propertyCellMapping_07_1, exls, out used07))
                //            return false;
                //    }
                //    else
                //    {
                //        if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use07, propertyCellMapping_07_2, exls, out used07))
                //            return false;
                //    }                        

                //    //08
                //    bool used08;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use08, propertyCellMapping_08, exls, out used08))
                //        return false;

                //    //09
                //    bool used09;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use09, propertyCellMapping_09, exls, out used09))
                //        return false;

                //    //10
                //    bool used10;
                //    if (!SetInput(process, dataAsIS, dataVariant, propertyCellMapping_Use10, propertyCellMapping_10, exls, out used10))
                //        return false;

                //    double? val = exls.GetCellValue(sheetOutput, kpiCellMapping[process.KpiId]) as double?;

                //    if (val == null)
                //        return false;

                //    kpiValue = Math.Round((double)val,1);
                //}
                


                //output = new Ecodistrict.Messaging.Data.Output(process.KpiId, kpiValue);

                //return true;

                #endregion

            }
            catch (System.Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "CalculateKpi", exception: ex);
                throw ex;
            }
        }

        private bool SetDistrictProperties(Dictionary<string,object> currentData, CExcel exls, Dictionary<string,string> propertyCellMapping, string sheetName)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                try
                {
                    if (currentData.ContainsKey(property.Key))
                    {
                        object value = currentData[property.Key];
                        Set(sheetName, property.Value, value, ref exls);
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    SendErrorMessage(
                        message:
                        String.Format(ex.Message + "\t key = {0}, CurrentDataMissing = {1}", property.Key,
                            currentData == null), sourceFunction: "SetDistrictProperties", exception: ex);
                    throw ex;
                }
            }
            return true;
        }

        bool SetInput(ModuleProcess process, Dictionary<string, object> dataAsIS, Dictionary<string, object> dataVariant, KeyValuePair<string, string> settingCellMapping, Dictionary<string, string> propertyCellMapping, CExcel exls, out bool Used)
        {
            //Used = true;
            //Set(sheetSettings, settingCellMapping.Value, "Y", ref exls);
            //if (!SetProperties(process, dataAsIS, exls, propertyCellMapping))
            //    return false;
            //return true;


            Used = false;

            if (dataVariant.ContainsKey(settingCellMapping.Key))
            {
                if (dataVariant[settingCellMapping.Key] as string == null)
                {
                    process.CalcMessage = "Module setting not properly set in database";
                    return false;
                }

                if (dataVariant[settingCellMapping.Key] as string == "Yes")
                {
                    Used = true;

                    Set(sheetSettings, settingCellMapping.Value, "Y", ref exls);

                    if (!SetProperties(process, dataAsIS, exls, propertyCellMapping))
                        return false;

                }
                else
                    exls.SetCellValue(sheetSettings, settingCellMapping.Value, "N");

            }
            else
                exls.SetCellValue(sheetSettings, settingCellMapping.Value, "N");

            return true;

        }

        private bool SetProperties(ModuleProcess process, Dictionary<string, object> data, CExcel exls, Dictionary<string, string> propertyCellMapping)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                try
                {
                    if (!CheckAndReportDistrictProp(process, data, property.Key))
                        return false;

                    object value = data[property.Key];

                    double val = Convert.ToDouble(value);

                    Set(sheetInput, property.Value, value, ref exls);
                }
                catch (System.Exception ex)
                {
                    SendErrorMessage(message: String.Format(ex.Message + "\t key = {0}, isDataMissing = {1}", property.Key, data == null), sourceFunction: "SetProperties", exception: ex);
                    throw ex;
                }
            }

            return true;
        }
        
    }
}


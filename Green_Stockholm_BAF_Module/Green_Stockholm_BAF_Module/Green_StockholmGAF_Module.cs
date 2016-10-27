using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace Green_StockholmGAF_Module
{
    class Green_StockholmGAF_Module : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_green = "green-area-factor";
        const string kpi_biodiversity = "biodiversity";
        const string kpi_social_value = "social-value";
        const string kpi_climate_adaptation = "climate-adaptation";
        const string sheet = "EXISTING";

        private const string inputResultName = "District input for Stockholm Green";

        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_green,                 "G68"},
            {kpi_biodiversity,          "F70"},
            {kpi_social_value,          "F71"},
            {kpi_climate_adaptation,    "F72"}
        };

        Dictionary<string, string> propertyCellMapping_STGreen = new Dictionary<string, string>()
        {
            {"greentotarea",                                                "F67"},            {"stgreenunsupportedgroundgreenery",                             "F5"},            {"stgreenplantbedgt800",                                         "F6"},            {"stgreenplantbedbetween600and800",                              "F7"},            {"stgreenplantbedbetween200and600",                              "F8"},            {"stgreengreenroofgt300",                                        "F9"},            {"stgreengreenroofbetween50and300",                              "F10"},            {"stgreengreeneryonwalls",                                       "F11"},            {"stgreenbalconyboxes",                                          "F12"},            {"stgreenwatersurfacepermanent",                                 "F47"},            {"stgreenopenhardsurfacesthatallowwatertogetthrough",            "F48"},            {"stgreengravelandsand",                                         "F49"},            {"stgreenconcreteslabswithjoints",                               "F50"},            {"stgreendiversityinthefieldlayer",                              "F14"},            {"stgreennaturalspeciesselection",                               "F15"},            {"stgreendiversityonthinesdumroofs",                             "F16"},            {"stgreenintegratedbalconyboxeswithclimbingplants",              "F17"},            {"stgreenbutterflyrestaurants",                                  "F18"},            {"stgreengeneralbushes",                                         "F19"},            {"stgreenberrybushes",                                           "F20"},            {"stgreenlargetrees",                                            "E21"},            {"stgreenmediumlargetrees",                                      "E22"},            {"stgreensmalltrees",                                            "E23"},            {"stgreenoaks",                                                  "E24"},            {"stgreenfruittrees",                                            "E25"},            {"stgreenfaunadepots",                                           "E26"},            {"stgreenbeetlefeeders",                                         "E27"},            {"stgreenbirdfeeders",                                           "E28"},            {"stgreenbiologicallyaccessiblepermanentwater",                  "F53"},            {"stgreendryareaswithplantsthattemporarilyfillwithrainwater",    "F54"},            {"stgreendelayofrainwaterinponds",                               "F55"},            {"stgreendelayofrainwaterinundergroundpercolationsystems",       "F56"},            {"stgreenrunofffromimpermeablesurfacestosurfaceswithplants",     "F57"},            {"stgreengrassareagames",                                        "F30"},            {"stgreengardeningareasinyards",                                 "F31"},            {"stgreenbalconiesandterracespreparedforgrowing",                "F32"},            {"stgreensharedroofterraces",                                    "F33"},            {"stgreenvisiblegreenroofs",                                     "F34"},            {"stgreenfloralarrangements",                                    "F35"},            {"stgreenexperimentalvaluesofbushes",                            "F36"},            {"stgreenberrybusheswithediblefruits",                           "F37"},            {"stgreentreesexperimentalvalue",                                "E38"},            {"stgreenfruittreesandbloomingtrees",                            "E39"},            {"stgreengreensurrounded",                                       "F40"},            {"stgreenbirdfeedersexperimentalvalue",                          "E41"},            {"stgreenwatersurfaces",                                         "F59"},            {"stgreenbiologicallyaccessiblewater",                           "F60"},            {"stgreenfountainscirculationssystems",                          "E61"},            {"stgreentreesleafyshading",                                     "E43"},            {"stgreenshadefromleafcover",                                    "F44"},            {"stgreeneveningoutoftemp",                                      "F45"},            {"stgreenwatercollectionduringdryperiods",                       "F63"},            {"stgreencollectedrainwaterforwatering",                         "F64"},            {"stgreenfountainscoolingeffect",                                "E65"}
        };


        #endregion

        #endregion

        public Green_StockholmGAF_Module()
        {
            this.useDummyDB = false;
            this.useBothVariantAndAsISForVariant = false;
            useXLSData = false;

            ////IMB-hub info (not used)
            //this.UserId = 0;
            //this.UserName = "";
            //this.ModuleName = "SP_Green_StockholmBAF_Module";

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
                    process.CalcMessage = "kpi not avaiable for this module";
                    return false;
                }
                
                if (process.CurrentData == null)
                {
                    process.CalcMessage = "Data missing";
                    return false;
                }

                if (!SetProperties(process, exls, propertyCellMapping_STGreen))
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
                //}

                //string outSheet = "EXISTING";
                //if (process.Request.variantId != null)
                //{
                //    outSheet = "PLANNED";
                //}
                //double? val = exls.GetCellValue(outSheet, kpiCellMapping[process.KpiId]) as double?;

                double? val = exls.GetCellValue(sheet, kpiCellMapping[process.KpiId]) as double?;
                
                if (val == null)
                    return false;

                double value = Math.Round((double)val, 2);

                //double value;
                //if (process.KpiId == kpi_green)
                //    value = Math.Round((double)val, 2);
                //else
                //    value = Math.Round((double)val * 100.0, 0);

                output = new Ecodistrict.Messaging.Data.Output(process.KpiId, value);

                return true;
            }
            catch (System.Exception ex)
            {
                SendErrorMessage(message: ex.Message, sourceFunction: "CalculateKpi", exception: ex);
                throw ex;
            }
        }

        private bool SetProperties(ModuleProcess process, CExcel exls, Dictionary<string, string> propertyCellMapping)
        {
            var nw = process.CurrentData[inputResultName] as List<object>;
            Dictionary<string, object> CurrentData = nw[0] as Dictionary<string, Object>;
            if (CurrentData == null)
                return false;

            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                try
                {
                    if (CurrentData.ContainsKey(property.Key))
                    {
                        object value = CurrentData[property.Key];
                        Set(sheet,property.Value,value,ref exls);
                    }
                    else
                    {
                        process.CalcMessage = "";
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
        
    }

}


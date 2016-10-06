using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;
using Ecodistrict.Messaging.Data;

namespace Green_BerlinBAF_Module
{
    class Green_BerlinBAF_Module : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_baf = "biotope-area-factor";
        const string kpi_berlin_baf = "berlin-biotope-area-factor";

        const string sheet = "Input";
        const string sheetOutput = "Input";

        private const string inputResultName = "District input for Berlin Green";

        Dictionary<string, InputSpecification> inputSpecifications;

        #region Cell Mapping

        private Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_berlin_baf, "D35"}
        };
 

        Dictionary<string,string> GreenCellMapping=new Dictionary<string, string>()
        {
            {"greentotarea",                "B4"},
            {"greensealedsurfacearea",      "C12"},
            {"greenpartsealedsurfacearea",  "C13"},
            {"greensemiopensurfacearea",    "C14"},
            {"greenvegetationlt80area",     "C15"},
            {"greenvegetationgt80area",     "C19"},
            {"greenvergetationtosoilarea",  "C23"},
            {"greenrainwaterinfarea",       "C25"},
            {"greenverticalarea",           "C27"},
            {"greenrooftoparea",            "C29"}
        };

        #endregion

        #endregion

        
        public Green_BerlinBAF_Module()
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
                        
        private void Set(string sheet, string cell, object value, ref CExcel exls)
        {
            if (!exls.SetCellValue(sheet, cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {2} in sheet {3}", cell, value, sheet));
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

        private bool SetProperties(ModuleProcess process, CExcel exls, Dictionary<string, string> propertyCellMapping)
        {
            var nw = process.CurrentData[inputResultName] as List<Object>;
            //var CurrentData = nw[0] as Dictionary<string, object>;

            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                Dictionary<string, object> CurrentData = nw[0] as Dictionary<string, object>;
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


        protected override bool CalculateKpi(ModuleProcess process, CExcel exls, out Output output, out OutputDetailed outputDetailed)
        {
            output = null;
            outputDetailed = null;

            if (!CheckAndReportDistrictProp(process, process.CurrentData, inputResultName))
                return false;

            if (!SetProperties(process, exls, GreenCellMapping))
                return false;

            double kpiValue;
            kpiValue = Convert.ToDouble(exls.GetCellValue(sheetOutput, kpiCellMapping[process.KpiId]));
            
            output=new Output(process.KpiId, Math.Round(kpiValue,1));
            return true;
        }        
    }
}


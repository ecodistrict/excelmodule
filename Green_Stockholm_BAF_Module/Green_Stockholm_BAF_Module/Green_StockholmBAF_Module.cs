using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace Green_StockholmBAF_Module
{
    class Green_StockholmBAF_Module : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_gaf = "green-area-factor";
        const string kpi_stockholm_gaf = "stockholm-green-area-factor";
        const string result_cell = "E30";

        Dictionary<string, InputSpecification> inputSpecifications;
        
        void DefineInputSpecifications()
        {
            try
            {
                inputSpecifications = new Dictionary<string, InputSpecification>();

                //GeoJson
                inputSpecifications.Add(kpi_gaf, GetInputSpecificationGreen());
                inputSpecifications.Add(kpi_stockholm_gaf, GetInputSpecificationGreen());
            }
            catch (System.Exception ex)
            {
                CExcelModule_ErrorRaised(this, ex);
            }
        }

        #endregion

        public Green_StockholmBAF_Module()
        {
            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "";
            this.ModuleName = "SP_Green_StockholmBAF_Module";

            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_gaf, kpi_stockholm_gaf };

            //Error handler
            this.ErrorRaised += CExcelModule_ErrorRaised;

            //Notification
            this.StatusMessage += CExcelModule_StatusMessage;
            
            //Define the input specification for the different kpis
            DefineInputSpecifications();
        }
                        
        private void Set(string sheet, string cell, object value, ref CExcel exls)
        {
            if (!exls.SetCellValue(sheet, cell, value))
                throw new Exception(String.Format("Could not set cell {} to value {2} in sheet {3}", cell, value, sheet));
        }

        string total_area = "total_area";
        string developed_area = "developed_area";
        string weighting_factor_per_m2 = "weighting_factor_per_m2";
        string amount = "amount";

        string area_info = "area_info";
        string sealed_surfaces = "sealed_surfaces";
        string partially_sealed_surfaces = "partially_sealed_surfaces";
        string semiOpenSurfacesStr = "semiOpenSurfaces";
        string swvutsBelow80Str = "swvutsBelow80";
        string swvutsAbove80Str = "swvutsAbove80";
        string swvConnectedToSoilBelowStr = "swvConnectedToSoilBelow";
        string rainwaterInfiltrationpSqrmRunoffAreaStr = "rainwaterInfiltrationpSqrmRunoffArea";
        string vgtm10mHeightStr = "vgtm10mHeight";
        string greeneryOnRooftopStr = "greeneryOnRooftop";

        InputSpecification GetInputSpecificationGreen()
        {
            InputSpecification iSpec = new InputSpecification();

            int ipgOrder = 0;
            InputGroup areaInfo = new InputGroup("Area information", ++ipgOrder);
            areaInfo.Add(total_area, new Number("Total area", 1, 0, "m\u00b2", 0));
            //areaInfo.Add(developed_area, new Number("Developed area", 2, 0, "m\u00b2", 0));
            iSpec.Add(area_info, areaInfo);

            //Sealed Surfaces
            InputGroup sealedSurfaces = new InputGroup("Sealed surfaces", ++ipgOrder);
            //sealedSurfaces.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0, null, 0));
            sealedSurfaces.Add(amount, new Number("Amount (weighted 0)", 2, 0, "m\u00b2", 0));
            iSpec.Add(sealed_surfaces, sealedSurfaces);

            //Partially Sealed Surfaces
            InputGroup partiallySealedSurfaces = new InputGroup("Partially sealed surfaces", ++ipgOrder);
            //partiallySealedSurfaces.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.3, null, 0));
            partiallySealedSurfaces.Add(amount, new Number("Amount (weighted 0.3)", 2, 0, "m\u00b2", 0));
            iSpec.Add(partially_sealed_surfaces, partiallySealedSurfaces);

            //Semi-open Surfaces
            InputGroup semiOpenSurfaces = new InputGroup("Semi-open surfaces", ++ipgOrder);
            //semiOpenSurfaces.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.5, null, 0));
            semiOpenSurfaces.Add(amount, new Number("Amount (weighted 0.5)", 2, 0, "m\u00b2", 0));
            iSpec.Add(semiOpenSurfacesStr, semiOpenSurfaces);

            //Surfaces with vegetation unconnected to the soil below and with < 80 mm of soil covering
            InputGroup swvutsBelow80 = new InputGroup("Surfaces with vegetation unconnected to the soil below and with < 80 mm of soil covering", ++ipgOrder);
            //swvutsBelow80.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.5, null, 0));
            swvutsBelow80.Add(amount, new Number("Amount (weighted 0.5)", 2, 0, "m\u00b2", 0));
            iSpec.Add(swvutsBelow80Str, swvutsBelow80);

            //Surfaces with vegetation unconnected to the soil below and with > 80 mm of soil covering
            InputGroup swvutsAbove80 = new InputGroup("Surfaces with vegetation unconnected to the soil below and with > 80 mm of soil covering", ++ipgOrder);
            //swvutsAbove80.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.7, null, 0));
            swvutsAbove80.Add(amount, new Number("Amount (weighted 0.7)", 2, 0, "m\u00b2", 0));
            iSpec.Add(swvutsAbove80Str, swvutsAbove80);

            //Surfaces with vegetation connected to the soil below
            InputGroup swvConnectedToSoilBelow = new InputGroup("Surfaces with vegetation connected to the soil below", ++ipgOrder);
            //swvConnectedToSoilBelow.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 1.0, null, 0));
            swvConnectedToSoilBelow.Add(amount, new Number("Amount (weighted 1.0)", 2, 0, "m\u00b2", 0));
            iSpec.Add(swvConnectedToSoilBelowStr, swvConnectedToSoilBelow);

            //Rainwater infiltration per m² of runoff area
            InputGroup rainwaterInfiltrationpSqrmRunoffArea = new InputGroup("Rainwater infiltration per m² of runoff area", ++ipgOrder);
            //rainwaterInfiltrationpSqrmRunoffArea.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.2, null, 0));
            rainwaterInfiltrationpSqrmRunoffArea.Add(amount, new Number("Amount (weighted 0.2)", 2, 0, "m\u00b2", 0));
            iSpec.Add(rainwaterInfiltrationpSqrmRunoffAreaStr, rainwaterInfiltrationpSqrmRunoffArea);

            //Vertical greenery up to a maximum of 10 m in height
            InputGroup vgtm10mHeight = new InputGroup("Vertical greenery up to a maximum of 10 m in height", ++ipgOrder);
            //vgtm10mHeight.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.5, null, 0));
            vgtm10mHeight.Add(amount, new Number("Amount (weighted 0.5)", 2, 0, "m\u00b2", 0));
            iSpec.Add(vgtm10mHeightStr, vgtm10mHeight);

            //Greenery on rooftop
            InputGroup greeneryOnRooftop = new InputGroup("Greenery on rooftop", ++ipgOrder);
            //greeneryOnRooftop.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.7, null, 0));
            greeneryOnRooftop.Add(amount, new Number("Amount (weighted 0.7)", 2, 0, "m\u00b2", 0));
            iSpec.Add(greeneryOnRooftopStr, greeneryOnRooftop);
            
            
            return iSpec;
        }

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            if (!inputSpecifications.ContainsKey(kpiId))
                throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.", kpiId));

            return inputSpecifications[kpiId];
        }

        protected override bool CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls, out Ecodistrict.Messaging.Output.Outputs outputs)
        {
            if (indata == null)
                throw new Exception("No data received!");

            outputs = new Ecodistrict.Messaging.Output.Outputs();
            
            #region Area Info
            {
                object value = 0;
                string key = area_info;
                if (indata.ContainsKey(key))
                {
                    if (indata[area_info] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[total_area] is Number)
                            value = (ipg[total_area] as Number).GetValue();
                    }

                }

                Set("Blad1", "B4", value, ref exls);
            }
            #endregion

            #region Area types data
            #region Sealed Surfaces
            {
                object value = 0;
                string key = sealed_surfaces;
                string cell = "C12";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion

            #region Partially Sealed Surfaces
            {
                object value = 0;
                string key = partially_sealed_surfaces;
                string cell = "C13";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion
            
            #region semiOpenSurfacesStr
            {
                object value = 0;
                string key = semiOpenSurfacesStr;
                string cell = "C14";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion
            
            #region swvutsBelow80Str
            {
                object value = 0;
                string key = swvutsBelow80Str;
                string cell = "C15";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion
            
            #region swvutsAbove80Str
            {
                object value = 0;
                string key = swvutsAbove80Str;
                string cell = "C19";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion
            
            #region swvConnectedToSoilBelowStr
            {
                object value = 0;
                string key = swvConnectedToSoilBelowStr;
                string cell = "C23";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion

            #region rainwaterInfiltrationpSqrmRunoffAreaStr
            {
                object value = 0;
                string key = rainwaterInfiltrationpSqrmRunoffAreaStr;
                string cell = "C25";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion

            #region vgtm10mHeightStr
            {
                object value = 0;
                string key = vgtm10mHeightStr;
                string cell = "C27";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion

            #region greeneryOnRooftopStr
            {
                object value = 0;
                string key = greeneryOnRooftopStr;
                string cell = "C29";
                if (indata.ContainsKey(key))
                {
                    if (indata[key] is InputGroup)
                    {
                        Dictionary<string, Input> ipg = (indata[key] as InputGroup).GetInputs();
                        if (ipg[amount] is Number)
                            value = (ipg[amount] as Number).GetValue();
                    }

                }

                Set("Blad1", cell, value, ref exls);
            }
            #endregion
            #endregion

            double kpi = Convert.ToDouble(exls.GetCellValue("Blad1", "D35"));

            switch (kpiId)
            {
                case kpi_gaf:
                case kpi_stockholm_gaf:
                    outputs.Add(new Ecodistrict.Messaging.Output.Kpi(Math.Round(kpi, 2), kpiId, ""));
                    break;
                default:
                    throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
            }


            return true;
        }

        protected override bool CalculateKpi(object indata, string kpiId, CExcel exls, out Ecodistrict.Messaging.Output.Outputs outputs)
        {
            outputs = new Ecodistrict.Messaging.Output.Outputs();
            return true;
        }
    }
}


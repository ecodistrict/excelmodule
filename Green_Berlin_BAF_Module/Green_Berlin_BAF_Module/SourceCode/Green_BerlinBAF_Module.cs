using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Ecodistrict.Messaging;
using Ecodistrict.Excel;

namespace Green_BerlinBAF_Module
{
    class Green_BerlinBAF_Module : CExcelModule
    {
        #region Defines
        // - Kpis
        const string kpi_berlin_baf = "biotope-area-factor";
        const string result_cell = "E30";

        Dictionary<string, InputSpecification> inputSpecifications;
        
        void DefineInputSpecifications()
        {
            try
            {
                inputSpecifications = new Dictionary<string, InputSpecification>();

                //GeoJson
                inputSpecifications.Add(kpi_berlin_baf, GetInputSpecificationGreen());
            }
            catch (System.Exception ex)
            {
                CExcelModule_ErrorRaised(this, ex);
            }
        }

        #endregion

        public Green_BerlinBAF_Module()
        {
            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "";
            this.ModuleName = "SP_Green_BerlinBAF_Module";

            //List of kpis the module can calculate
            this.KpiList = new List<string> { kpi_berlin_baf };

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
            areaInfo.Add(developed_area, new Number("Developed area", 2, 0, "m\u00b2", 0));
            iSpec.Add(area_info, areaInfo);

            //Sealed Surfaces
            InputGroup sealedSurfaces = new InputGroup("Sealed surfaces", ++ipgOrder);
            sealedSurfaces.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0, null, 0));
            sealedSurfaces.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(sealed_surfaces, sealedSurfaces);

            //Partially Sealed Surfaces
            InputGroup partiallySealedSurfaces = new InputGroup("Partially sealed surfaces", ++ipgOrder);
            partiallySealedSurfaces.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.3, null, 0));
            partiallySealedSurfaces.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(partially_sealed_surfaces, partiallySealedSurfaces);

            //Semi-open Surfaces
            InputGroup semiOpenSurfaces = new InputGroup("Semi-open surfaces", ++ipgOrder);
            semiOpenSurfaces.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.5, null, 0));
            semiOpenSurfaces.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(semiOpenSurfacesStr, semiOpenSurfaces);

            //Surfaces with vegetation unconnected to the soil below and with < 80 mm of soil covering
            InputGroup swvutsBelow80 = new InputGroup("Surfaces with vegetation unconnected to the soil below and with < 80 mm of soil covering", ++ipgOrder);
            swvutsBelow80.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.5, null, 0));
            swvutsBelow80.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(swvutsBelow80Str, swvutsBelow80);

            //Surfaces with vegetation unconnected to the soil below and with > 80 mm of soil covering
            InputGroup swvutsAbove80 = new InputGroup("Surfaces with vegetation unconnected to the soil below and with > 80 mm of soil covering", ++ipgOrder);
            swvutsAbove80.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.7, null, 0));
            swvutsAbove80.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(swvutsAbove80Str, swvutsAbove80);

            //Surfaces with vegetation connected to the soil below
            InputGroup swvConnectedToSoilBelow = new InputGroup("Surfaces with vegetation connected to the soil below", ++ipgOrder);
            swvConnectedToSoilBelow.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 1.0, null, 0));
            swvConnectedToSoilBelow.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(swvConnectedToSoilBelowStr, swvConnectedToSoilBelow);

            //Rainwater infiltration per m² of runoff area
            InputGroup rainwaterInfiltrationpSqrmRunoffArea = new InputGroup("Rainwater infiltration per m² of runoff area", ++ipgOrder);
            rainwaterInfiltrationpSqrmRunoffArea.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.2, null, 0));
            rainwaterInfiltrationpSqrmRunoffArea.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(rainwaterInfiltrationpSqrmRunoffAreaStr, rainwaterInfiltrationpSqrmRunoffArea);

            //Vertical greenery up to a maximum of 10 m in height
            InputGroup vgtm10mHeight = new InputGroup("Vertical greenery up to a maximum of 10 m in height", ++ipgOrder);
            vgtm10mHeight.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.5, null, 0));
            vgtm10mHeight.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(vgtm10mHeightStr, vgtm10mHeight);

            //Greenery on rooftop
            InputGroup greeneryOnRooftop = new InputGroup("Greenery on rooftop", ++ipgOrder);
            greeneryOnRooftop.Add(weighting_factor_per_m2, new Number("Weighting factor per m²", 1, 0.7, null, 0));
            greeneryOnRooftop.Add(amount, new Number("Amount", 2, 0, "m\u00b2", 0));
            iSpec.Add(greeneryOnRooftopStr, greeneryOnRooftop);
            
            
            return iSpec;
        }

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            if (!inputSpecifications.ContainsKey(kpiId))
                throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.", kpiId));

            return inputSpecifications[kpiId];
        }

        protected override Ecodistrict.Messaging.Output.Outputs CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls)
        {
            Ecodistrict.Messaging.Output.Outputs outputs = new Ecodistrict.Messaging.Output.Outputs();
            
            #region Area Info
            if (indata[area_info] is InputGroup)
            {
                Dictionary<string,Input> ipg = (indata[area_info] as InputGroup).GetInputs();
                if (ipg[total_area] is Number)
                    Set("Blad1", "B4", (ipg[total_area] as Number).GetValue(), ref exls);

                if (ipg[developed_area] is Number)
                    Set("Blad1", "C4", (ipg[developed_area] as Number).GetValue(), ref exls);
            }
            #endregion

            #region Sealed Surfaces
            if (indata[sealed_surfaces] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[sealed_surfaces] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B12", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C12", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion

            #region Partially Sealed Surfaces
            if (indata[partially_sealed_surfaces] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[partially_sealed_surfaces] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B13", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C13", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion
            
            #region semiOpenSurfacesStr
            if (indata[semiOpenSurfacesStr] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[semiOpenSurfacesStr] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B14", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C14", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion
            
            #region swvutsBelow80Str
            if (indata[swvutsBelow80Str] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[swvutsBelow80Str] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B15", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C15", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion
            
            #region swvutsAbove80Str
            if (indata[swvutsAbove80Str] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[swvutsAbove80Str] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B19", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C19", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion
            
            #region swvConnectedToSoilBelowStr
            if (indata[swvConnectedToSoilBelowStr] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[swvConnectedToSoilBelowStr] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B23", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C23", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion

            #region rainwaterInfiltrationpSqrmRunoffAreaStr
            if (indata[rainwaterInfiltrationpSqrmRunoffAreaStr] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[rainwaterInfiltrationpSqrmRunoffAreaStr] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B25", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C25", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion

            #region vgtm10mHeightStr
            if (indata[vgtm10mHeightStr] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[vgtm10mHeightStr] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B27", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C27", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion

            #region greeneryOnRooftopStr
            if (indata[greeneryOnRooftopStr] is InputGroup)
            {
                Dictionary<string, Input> ipg = (indata[greeneryOnRooftopStr] as InputGroup).GetInputs();
                if (ipg[weighting_factor_per_m2] is Number)
                    Set("Blad1", "B29", (ipg[weighting_factor_per_m2] as Number).GetValue(), ref exls);

                if (ipg[amount] is Number)
                    Set("Blad1", "C29", (ipg[amount] as Number).GetValue(), ref exls);
            }
            #endregion

            double kpi = Convert.ToDouble(exls.GetCellValue("Blad1", "D35"));

            switch (kpiId)
            {
                case kpi_berlin_baf:
                    outputs.Add(new Ecodistrict.Messaging.Output.Kpi(Math.Round(kpi, 2), kpi_berlin_baf, ""));
                    break;
                default:
                    throw new ApplicationException(String.Format("No calculation procedure could be found for '{0}'", kpiId));
            }


            return outputs;
        }

        public bool Init(string IMB_config_path, string Module_config_path)
        {
            try
            {
                Init_IMB(IMB_config_path);
                Init_Module(Module_config_path);
                return true;
            }
            catch (Exception ex)
            {
                CExcelModule_ErrorRaised(this, ex);
                return false;
            }
        }

    }
}


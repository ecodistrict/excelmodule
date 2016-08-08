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

        #region Cell Mapping
        Dictionary<string, string> kpiCellMapping = new Dictionary<string, string>()
        {
            {kpi_green,                 "G68"},
            {kpi_biodiversity,          "F70"},
            {kpi_social_value,          "F71"},
            {kpi_climate_adaptation,    "F72"}
        };

        Dictionary<string, string> propertyCellMapping_Green = new Dictionary<string, string>()
        {
            {"Total land area",  "F67"}
        };


        Dictionary<string, string> propertyCellMapping_BSK = new Dictionary<string, string>()
        {
            {"Unsupported ground greenery",     "F5"},                              //Sub-factors greenery
            {"Plant bed (>800 mm)",             "F6"},
            {"Plant bed (600 - 800 mm)",        "F7"},
            {"Plant bed (200 - 600 mm)",        "F8"},
            {"Green roof (>300 mm)",            "F9"},
            {"Green roof (50 - 300 mm)",        "F10"},
            {"Greenery on walls",               "F11"},
            {"Balcony boxes",                   "F12"},

            {"Water surface permanent",                             "F47"},         //Sub-factors water
            {"Open hard surfaces that allow water to get through",  "F48"},
            {"Gravel and sand",                                     "F49"}
        };

        Dictionary<string, string> propertyCellMapping_SK = new Dictionary<string, string>()
        {                                
            {"Concrete slabs with joints",  "F50"}                                  //Sub-factors water
        };

        Dictionary<string, string> propertyCellMapping_B = new Dictionary<string, string>()
        {
            {"Diversity in the field layer",                    "F14"},             //Supplementary factors greenery/biodiversity
            {"Natural species selection",                       "F15"},
            {"Diversity on thin sedum roofs",                   "F16"},
            {"Integrated balcony boxes with climbing plants",   "F17"},
            {"Butterfly restaurants",                           "F18"},
            {"General bushes",                                  "F19"},
            {"Berry bushes",                                    "F20"},
            {"Large trees",                                     "E21"},
            {"Medium large trees",                              "E22"},
            {"Small trees",                                     "E23"},
            {"Oaks",                                            "E24"},
            {"Fruit trees",                                     "E25"},
            {"Fauna depots",                                    "E26"},
            {"Beetle feeders",                                  "E27"},
            {"Bird feeders",                                    "E28"},

            {"Biologically accessible permanent water",                     "F53"},             //Supplementary factors water/biodiversity
            {"Dry areas with plants that temporarily fill with rain water", "F54"},
            {"Delay of rainwater in ponds",                                 "F55"},
            {"Delay of rainwater in underground percolation systems",       "F56"},
            {"Runoff from impermeable surfaces to surfaces with plants",    "F57"}
        };

        Dictionary<string, string> propertyCellMapping_S = new Dictionary<string, string>()
        {
            {"Grass area games",                                        "F30"},     //Supplementary factors greenery/recreation and social value
            {"Gardening areas in yards",                                "F31"},
            {"Balconies and terraces prepared for growing",             "F32"},
            {"Shared roof terraces",                                    "F33"},
            {"Visible green roofs",                                     "F34"},
            {"Floral arrangements",                                     "F35"},
            {"Experiential values of bushes",                           "F36"},
            {"Berry bushes with edible fruits",                         "F37"},
            {"Trees experiential value",                                "E38"},
            {"Fruit trees and blooming trees",                          "E39"},
            {"Green surrounded",                                        "F40"},
            {"Bird feeders experiential value",                         "E41"},

            {"Water surfaces",                  "F59"},         //Supplementary factors water/recreational and social values
            {"Biologically accessible water",   "F60"},
            {"Fountains circulations systems",  "E61"}
        };

        Dictionary<string, string> propertyCellMapping_K = new Dictionary<string, string>()
        {
            {"Trees leafy shading",     "E43"}, //Supplementary factors greenery/climate - heat islands
            {"Shade from leaf cover",   "F44"},
            {"Evening out of temp",     "F45"},
            
            {"Water collection during dry periods", "F63"}, //Supplementary factors water/climate - heat islands
            {"Collected rainwater for watering",    "F64"},
            {"Fountains cooling effect",            "E65"}
        };

        #region Old Style

        Dictionary<string, string> propertyCellMapping_Green_Old = new Dictionary<string, string>()
        {
            {"total_area",  "F67"}
        };


        Dictionary<string, string> propertyCellMapping_BSK_Old = new Dictionary<string, string>()
        {
            {"unsopported_ground_greenery",     "F5"},                              //Sub-factors greenery
            {"plant_bed_above_800mm",           "F6"},
            {"plant_bed_between_600and800mm",   "F7"},
            {"plant_bed_between_200and600mm",   "F8"},
            {"green_roof_above_300mm",          "F9"},
            {"green_roof_between_50and300mm",   "F10"},
            {"greenery_on_walls",               "F11"},
            {"balcony_boxes",                   "F12"},

            {"water_surface_permanent",                             "F47"},         //Sub-factors water
            {"open_hard_surfaces_that_allow_water_to_get_through",  "F48"},
            {"gravel_and_sand",                                     "F49"}
        };

        Dictionary<string, string> propertyCellMapping_SK_Old = new Dictionary<string, string>()
        {                                
            {"concrete_slabs_with_joints",  "F50"}                                  //Sub-factors water
        };

        Dictionary<string, string> propertyCellMapping_B_Old = new Dictionary<string, string>()
        {
            {"diversity_in_field_layer",                        "F14"},             //Supplementary factors greenery/biodiversity
            {"natural_species_selection",                       "F15"},
            {"diversity_on_thin_sedum_roofs",                   "F16"},
            {"integrated_balcony_boxes_with_climbing_plants",   "F17"},
            {"butterfly_restaurants",                           "F18"},
            {"bushes_general",                                  "F19"},
            {"berry_bushes",                                    "F20"},
            {"large_trees_trunkabove_30cm",                     "E21"},
            {"medium_large_trees_trunk_between_20and30cm",      "E22"},
            {"small_trees_trunk_between_16and20cm",             "E23"},
            {"oak_quercus_robur",                               "E24"},
            {"fruit_trees",                                     "E25"},
            {"fauna_depots",                                    "E26"},
            {"beetle_feeders",                                  "E27"},
            {"bird_feeders",                                    "E28"},

            {"biologically_accessible_permanent_water",                     "F53"},             //Supplementary factors water/biodiversity
            {"dry_areas_with_plants_that_temporarily_fill_with_rainwater",  "F54"},
            {"delay_of_rainwater_in_ponds",                            "F55"},
            {"delay_of_rainwater_in_underground_percolation_systems",       "F56"},
            {"runoff_from_impermeable_surfaces_to_surfaces_with_plants",    "F57"}
        };

        Dictionary<string, string> propertyCellMapping_S_Old = new Dictionary<string, string>()
        {
            {"grass_area_usable_for_ball_games_and_playing",            "F30"},     //Supplementary factors greenery/recreation and social value
            {"gardening_areas_in_yards",                                "F31"},
            {"balconies_and_terraces_prepared_for_growing",             "F32"},
            {"shared_roof_terraces",                                    "F33"},
            {"visible_green_roofs",                                     "F34"},
            {"Floral arrangements",                                     "F35"},
            {"experiential_value_of_bushes",                            "F36"},
            {"berry_bushes_with_edible_fruit_etc)",                     "F37"},
            {"trees_experiential_value",                                "E38"},
            {"fruit_trees_and_blooming_trees",                          "E39"},
            {"pergolas_paths_surrounded_by_leaves_and_other_greenery",  "F40"},
            {"bird_feeders_experiential_value",                         "E41"},

            {"water_surfaces",                                      "F59"},         //Supplementary factors water/recreational and social values
            {"biologically_accessible_water_experiential_value",    "F60"},
            {"fountains_circulations_systems_etc",                  "E61"}
        };

        Dictionary<string, string> propertyCellMapping_K_Old = new Dictionary<string, string>()
        {
            {"trees_with_leafy_shade_over_play_areas_etc",                  "E43"}, //Supplementary factors greenery/climate - heat islands
            {"pergolas_green_corridors_etc_equals_shade_from_leaf_cover",   "F44"},
            {"green_roofs_ground_greenery_evening_out_of_temp",            "F45"},
            
            {"water_collection_during_dry_periods",                 "F63"}, //Supplementary factors water/climate - heat islands
            {"collected_rainwater_for_watering_climate_impact",     "F64"},
            {"fountains_etc_cooling_effect",                        "E65"}
        };

        #endregion

        #endregion

        #endregion

        public Green_StockholmGAF_Module()
        {
            useDummyDB = false;
            useXLSData = false;

            //IMB-hub info (not used)
            this.UserId = 0;
            this.UserName = "";
            this.ModuleName = "SP_Green_StockholmBAF_Module";

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

        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            var iSpec = new InputSpecification();

            if (!KpiList.Contains(kpiId))
                return null;

            //SetIspec(ref iSpec, propertyCellMapping_Green);
            //SetIspec(ref iSpec, propertyCellMapping_BSK);
            //SetIspec(ref iSpec, propertyCellMapping_SK);
            //SetIspec(ref iSpec, propertyCellMapping_B);
            //SetIspec(ref iSpec, propertyCellMapping_S);
            //SetIspec(ref iSpec, propertyCellMapping_K);

            if (kpiId == kpi_green)
                SetIspec(ref iSpec, propertyCellMapping_Green_Old);

            if ((kpiId == kpi_green) |
                (kpiId == kpi_biodiversity) |
                (kpiId == kpi_social_value) |
                (kpiId == kpi_climate_adaptation))
                SetIspec(ref iSpec, propertyCellMapping_BSK_Old);

            if ((kpiId == kpi_green) |
                (kpiId == kpi_social_value) |
                (kpiId == kpi_climate_adaptation))
                SetIspec(ref iSpec, propertyCellMapping_SK_Old);

            if ((kpiId == kpi_green) |
                (kpiId == kpi_biodiversity))
                SetIspec(ref iSpec, propertyCellMapping_B_Old);

            if ((kpiId == kpi_green) |
                (kpiId == kpi_social_value))
                SetIspec(ref iSpec, propertyCellMapping_S_Old);

            if ((kpiId == kpi_green) |
                (kpiId == kpi_climate_adaptation))
                SetIspec(ref iSpec, propertyCellMapping_K_Old);

            return iSpec;

        }

        void SetIspec(ref InputSpecification iSpec, Dictionary<string, string> propertyCellMapping)
        {
            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                if (!iSpec.ContainsKey(property.Key))
                    iSpec.Add(property.Key, new Number(property.Key));
            }
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
                
                if (process.KpiId == kpi_green)
                    if (!SetProperties(process, exls, propertyCellMapping_Green))
                        return false;

                if ((process.KpiId == kpi_green) |
                    (process.KpiId == kpi_biodiversity) |
                    (process.KpiId == kpi_social_value) |
                    (process.KpiId == kpi_climate_adaptation))
                    if (!SetProperties(process, exls, propertyCellMapping_BSK))
                        return false;

                if ((process.KpiId == kpi_green) |
                    (process.KpiId == kpi_social_value) |
                    (process.KpiId == kpi_climate_adaptation))
                    if (!SetProperties(process, exls, propertyCellMapping_SK))
                        return false;

                if ((process.KpiId == kpi_green) |
                    (process.KpiId == kpi_biodiversity))
                    if (!SetProperties(process, exls, propertyCellMapping_B))
                        return false;

                if ((process.KpiId == kpi_green) |
                    (process.KpiId == kpi_social_value))
                    if (!SetProperties(process, exls, propertyCellMapping_S))
                        return false;

                if ((process.KpiId == kpi_green) |
                    (process.KpiId == kpi_climate_adaptation))
                    if (!SetProperties(process, exls, propertyCellMapping_K))
                        return false;
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
            Dictionary<string, object> CurrentData = process.CurrentData;  

            foreach (KeyValuePair<string, string> property in propertyCellMapping)
            {
                try
                {
                    {
                        if (!CheckAndReportDistrictProp(process,CurrentData, property.Key))
                            return false;

                        object value = CurrentData[property.Key];

                        double val = Convert.ToDouble(value);
                        if (val < 0)
                        {
                            process.CalcMessage = String.Format("Property '{0}' has invalid data, only values equal or above zero is allowed; value: {1}", property.Key, val);
                            return false;
                        }

                        Set(sheet, property.Value, value, ref exls);
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


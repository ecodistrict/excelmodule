using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ecodistrict.Excel;
using Ecodistrict.Messaging;

namespace Cheese_Module
{
    public class CheeseModule : CExcelModule
    {
        private const string kpi_CheeseTaste = "cheese-taste-kpi";
        private const string kpi_CheesePrice = "cheese-price-kpi";
        private Dictionary<string, InputSpecification> inputSpecifications;


        /// <summary>
        /// Constructor
        /// </summary>
        public CheeseModule()
        {
            //Variables not used for now.
            this.UserId = 0;
            this.UserName = "";
            //Kpi list
            this.KpiList = new List<string> {kpi_CheeseTaste, kpi_CheesePrice};

            //Error handling
            this.ErrorRaised += CExcelModule_ErrorRaised;

            //Notifications
            this.StatusMessage += CExcelModule_StatusMessage;

            //Define the input specification for the different kpis
            DefineInputSpecifications();
        }

        /// <summary>
        /// Setts communcation and Excel document variables
        /// </summary>
        internal void Init()
        {
            //In this sample we use hardcoded variables
            this.RemoteHost = "vps17642.public.cloudvps.com";
            this.SubScribedEventName = "modules";
            this.PublishedEventName = "dashboard";

            this.ModuleName = "Cheese Module";
            this.Description = "A module to access cheese quality";
            this.ModuleId = "foo-bar_cheese-module-v1-1";
            //moduleSettings.path = @"C:\ECODistr-ICT\Exceldocuments\EcoCheeseSample.xlsx";
            this.WorkBookPath = "EcoCheeseSample.xlsx";
        }

        private void DefineInputSpecifications()
        {
            try
            {
                inputSpecifications = new Dictionary<string, InputSpecification>();
                inputSpecifications.Add(kpi_CheeseTaste, GetInputSpecification_CheeseTaste());
                inputSpecifications.Add(kpi_CheesePrice, GetInputSpecification_CheesePrice());
            }
            catch (Exception ex)
            {
                var exNew = new ErrorMessageEventArg()
                {
                    Exception = ex,
                    Message = ex.Message,
                    SourceFunction = "DefineInputSpecification"
                };
                CExcelModule_ErrorRaised(this, exNew);
            }
        }

        private InputSpecification GetInputSpecification_CheesePrice()
        {
            try
            {
                InputSpecification priceInputSpecification = new InputSpecification();
                priceInputSpecification.Add("Cheddar_Price",
                    new Number(label: "Cheddar price index", order: 0, min: 0.0, max: 15.0));
                priceInputSpecification.Add("Gamle_Ole_Price",
                    new Number(label: "Gamle Ole price index", order: 1, min: 0.0, max: 15.0));
                priceInputSpecification.Add("Vasterbotten_Price",
                    new Number(label: "Vasterbotten price index", order: 2, min: 0.0, max: 15.0));
                priceInputSpecification.Add("Edamer_Price",
                    new Number(label: "Edamer price index", order: 3, min: 0.0, max: 15.0));
                priceInputSpecification.Add("Maasdamer_Price",
                    new Number(label: "Maasmer price index", order: 4, min: 0.0, max: 15.0));
                priceInputSpecification.Add("Gouda_Price",
                    new Number(label: "Gouda price index", order: 5, min: 0.0, max: 15.0));
                return priceInputSpecification;
            }
            catch (Exception ex)
            {
                var exNew = new ErrorMessageEventArg()
                {
                    Exception = ex,
                    SourceFunction = "GetInputSpecification_CheesePrice",
                    Message = "Could not create input specification for cheese price!"
                };
                CExcelModule_ErrorRaised(this, exNew);
                return new InputSpecification();
            }
        }

        private InputSpecification GetInputSpecification_CheeseTaste()
        {
            try
            {
                InputSpecification tasteInputSpecification = new InputSpecification();
                tasteInputSpecification.Add("Cheddar_Taste",
                    new Number(label: "Cheddar taste index", order: 0, min: 0.0, max: 1.0));
                tasteInputSpecification.Add("Gamle_Ole_Taste",
                    new Number(label: "Gamle Ole taste index", order: 1, min: 0.0, max: 1.0));
                tasteInputSpecification.Add("Vasterbotten_Taste",
                    new Number(label: "Vasterbotten taste index", order: 2, min: 0.0, max: 1.0));
                tasteInputSpecification.Add("Edamer_Taste",
                    new Number(label: "Edamer taste index", order: 3, min: 0.0, max: 1.0));
                tasteInputSpecification.Add("Maasdamer_Taste",
                    new Number(label: "Maasmer taste index", order: 4, min: 0.0, max: 1.0));
                tasteInputSpecification.Add("Gouda_Taste",
                    new Number(label: "Gouda taste index", order: 5, min: 0.0, max: 1.0));
                return tasteInputSpecification;

            }
            catch (Exception ex)
            {
                var exNew = new ErrorMessageEventArg()
                {
                    Exception = ex,
                    SourceFunction = "GetInputSpecification_CheeseTaste",
                    Message = "Could not create input specification for cheese taste!"
                };
                CExcelModule_ErrorRaised(this, exNew);
                return new InputSpecification();
            }
        }

        /// <summary>
        /// This function has to be overrided from the base class. 
        /// Dependent of the KpiId it returns the correct input specification as a InputSpecification object 
        /// </summary>
        /// <param name="kpiId">The Kpi id of a Kpi in the module</param>
        /// <returns>Input specification for the kpiId</returns>
        protected override InputSpecification GetInputSpecification(string kpiId)
        {
            if (!inputSpecifications.ContainsKey(kpiId))
                throw new ApplicationException(String.Format("No input specification for kpiId '{0}' could be found.",
                    kpiId));

            return inputSpecifications[kpiId];
        }

        /// <summary>
        /// Calculates the Kpi when the dashboard has sent a StartModuleRequest
        /// </summary>
        /// <param name="indata">Collection of indata as a dictionary. Key is the name of the parameter
        /// and Value is an object variable with value of the parameter</param>
        /// <param name="kpiId">Kpi id. Decides which Kpi should be calculated</param>
        /// <param name="exls">Excel object in which the calculations should take place</param>
        /// <returns></returns>
        protected override Ecodistrict.Messaging.Output.Outputs CalculateKpi(Dictionary<string, Input> indata, string kpiId, CExcel exls)
        {
            Ecodistrict.Messaging.Output.Outputs outputs;
            switch (kpiId)
            {
                case kpi_CheeseTaste:
                    outputs = CalcCheeseTasteKpi(indata, exls);
                    break;
                case kpi_CheesePrice:
                    outputs = CalcCheesePriceKpi(indata, exls);
                    break;
                default:
                    throw new ArgumentException(string.Format("Kpi id unknown! ({0})", kpiId));
            }

            return outputs;
        }

        private Ecodistrict.Messaging.Output.Outputs CalcCheesePriceKpi(Dictionary<string, Input> indata, CExcel exls)
        {
            Ecodistrict.Messaging.Output.Outputs outputs = new Ecodistrict.Messaging.Output.Outputs();

            try
            {
                foreach (var input in indata)
                {
                    //Set all input values
                    switch (input.Key)
                    {
                        case "Cheddar_Price": //C3 cell
                            exls.SetCellValue("Sheet1", "C3", input.Value);
                            break;
                        case "Gamle_Ole_Price":
                            exls.SetCellValue("Sheet1", "C4", input.Value);
                            break;
                        case "Vasterbotten_Price":
                            exls.SetCellValue("Sheet1", "C5", input.Value);
                            break;
                        case "Edamer_Price":
                            exls.SetCellValue("Sheet1", "C6", input.Value);
                            break;
                        case "Maasdamer_Price":
                            exls.SetCellValue("Sheet1", "C7", input.Value);
                            break;
                        case "Gouda_Price":
                            exls.SetCellValue("Sheet1", "C8", input.Value);
                            break;
                        default:
                            break;
                    }
                }

                //Get the Kpi value
                var kpiValue = exls.GetCellValue("Sheet1", "C17");

                //Put it in the calc result
                outputs.Add(new Ecodistrict.Messaging.Output.Kpi(kpiValue, "Min Cheese price Kpi", "SEK"));

                return outputs;

            }
            catch (Exception ex)
            {

                throw new Exception(string.Format("Could not calculate the cheese price Kpi!\n{0}", ex));
            }
        }

        private Ecodistrict.Messaging.Output.Outputs CalcCheeseTasteKpi(Dictionary<string, Input> indata, CExcel exls)
        {
            Ecodistrict.Messaging.Output.Outputs outputs = new Ecodistrict.Messaging.Output.Outputs();

            try
            {
                foreach (var input in indata)
                {
                    //Set all input values
                    switch (input.Key)
                    {
                        case "Cheddar_Taste": //C3 cell
                            exls.SetCellValue("Sheet1", "B3", input.Value);
                            break;
                        case "Gamle_Ole_Taste":
                            exls.SetCellValue("Sheet1", "B4", input.Value);
                            break;
                        case "Vasterbotten_Taste":
                            exls.SetCellValue("Sheet1", "B5", input.Value);
                            break;
                        case "Edamer_Taste":
                            exls.SetCellValue("Sheet1", "B6", input.Value);
                            break;
                        case "Maasdamer_Taste":
                            exls.SetCellValue("Sheet1", "B7", input.Value);
                            break;
                        case "Gouda_Taste":
                            exls.SetCellValue("Sheet1", "B8", input.Value);
                            break;
                        default:
                            break;
                    }
                }

                //Get the Kpi value
                var kpiValue = exls.GetCellValue("Sheet1", "B17");

                //Put it in the calc result
                outputs.Add(new Ecodistrict.Messaging.Output.Kpi(kpiValue, "Max Cheese taste Kpi", "Unit"));
                return outputs;

            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Could not calculate the cheese taste Kpi!\n{0}", ex));
            }
        }
    }
}

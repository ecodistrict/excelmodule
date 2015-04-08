using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel=Microsoft.Office.Interop.Excel;


namespace EcoExcel
{
    /// <summary>
    /// Class that handles interaction with Microsoft Excel
    /// </summary>
    /// <remarks>
    /// This class opens the workbook referenced in the constructor in readonly mode.<br/>
    /// The Excelfile can be in Excel 97-2003 format (*.xls) or Later (*.xlsx)<br/>
    /// The class uses Microsoft.Office.Interop.Excel which means that Excel is started as a com object.
    /// </remarks>
    public class CExcel 
    {
        /// <summary>
        /// /complete path to the excel document that shuld be used, string
        /// </summary>
        private string WorkBookFilePath { get; set; }
        /// <summary>
        /// The Excel application object, com object
        /// </summary>
        private Excel.Application excelApp;
        /// <summary>
        /// Excel workbook object
        /// </summary>
        private Excel.Workbook _wb;
        /// <summary>
        /// Excel worksheet object
        /// </summary>
        private Excel.Worksheet _ws;



        /// <summary>
        /// Constructor. Opens the workbook related in the parameter workbookFilePath in readonly mode.
        /// </summary>
        /// <param name="workBookFilePath">The complete path to the workbook the class will use. </param>
        public CExcel(string workBookFilePath)
        {
            WorkBookFilePath = workBookFilePath;
            
            try
            {
                excelApp = new Excel.Application();
                _wb = excelApp.Workbooks.Open(WorkBookFilePath,ReadOnly:true);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Destructor
        /// </summary>
        /// <remarks>
        /// Closes down Excel if it is not already closed. 
        /// Uses the same code as <see cref="CloseExcel"/>
        /// </remarks>
        ~CExcel()
        {
            CloseExcel();
        }

        /// <summary>
        /// Sets the value using the specified parameters, using row and column syntax 
        /// </summary>
        /// <param name="sheet">Sheetname, string</param>
        /// <param name="row">Row, integer</param>
        /// <param name="col">Column, integer</param>
        /// <param name="value">Value, Object</param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        public bool SetCellValue(string sheet, int row, int col,object value)
        {
            try
            {
                if (_ws == null || _ws.Name != sheet)
                    _ws = _wb.Worksheets[sheet];

                _ws.Range[row, col].Value = value;
                return true;
            }
            catch (Exception)
            {
                
                return false;
            }
        }

        /// <summary>
        /// Sets the value using the specified parameters, using A1, A2, B1 syntax.
        /// </summary>
        /// <param name="sheet">Sheetname, string</param>
        /// <param name="cell">Cell, string (A1, A2, B1 syntax)</param>
        /// <param name="value">Value, object</param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        public bool SetCellValue(string sheet, string cell, object value)
        {
            try
            {
                if (_ws == null || _ws.Name != sheet)
                    _ws = _wb.Worksheets[sheet];

                _ws.Range[cell].Value = value;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Returns a cell value as object, using A1, A2, B1 syntax.
        /// </summary>
        /// <remarks>Throws an exception if operation fails</remarks>
        /// <param name="sheet">Sheetname, string</param>
        /// <param name="cell">Cell, string (A1, A2, B1 syntax)</param>
        /// <returns>Cellvalue as object</returns>
        public object GetCellValue(string sheet, string cell)
        {
            try
            {
                if (_ws == null || _ws.Name != sheet)
                    _ws = _wb.Worksheets[sheet];

                //celltype = ws.Range[cell].Value.GetType();
                return _ws.Range[cell].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} cell:{1}" , sheet, cell));
            }
        }

        /// <summary>
        /// Returns a cell value as object, using row and column syntax
        /// </summary>
        /// <remarks>Throws an exception if operation fails</remarks>
        /// <param name="sheet">Sheetname, string</param>
        /// <param name="row">Row, integer</param>
        /// <param name="col">Column, integer</param>
        /// <returns>Cellvalue as object</returns>
        public object GetCellValue(string sheet, int row, int col)
        {
            try
            {
                if (_ws == null || _ws.Name != sheet)
                    _ws = _wb.Worksheets[sheet];

                //celltype = ws.Range[row, col].Value.GetType();
                return _ws.Range[row, col].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} row:{1} col:{2}", sheet, row, col));
            }
        }

        /// <summary>
        /// Close the workbook and close the excel com object.
        /// </summary>
        /// <remarks>
        /// IMPORTANT!<br/>
        /// This routine should be used to ensure that excel closes down.
        /// </remarks>
        public void CloseExcel()
        {
            GC.Collect();
            if (_ws != null)
            {
                try
                {
                    excelApp.DisplayAlerts = false;
                    Marshal.FinalReleaseComObject(_ws);
                    _ws = null;
                    GC.Collect();
                }
                catch (Exception)
                {   
                }
            }
            if (_wb != null)
            {
                try
                {
                    excelApp.DisplayAlerts = false;
                    _wb.Close(false, Type.Missing, Type.Missing);
                    Marshal.FinalReleaseComObject(_wb);
                    _wb = null;
                    GC.Collect();
                }
                catch (Exception)
                {                    
                }
            }
            //if (excelApp.Workbooks != null)
            //{
            //    try
            //    {
            //        excelApp.Workbooks.Close();
            //        Marshal.FinalReleaseComObject(excelApp.Workbooks);
            //        GC.Collect();
            //    }
            //    catch (Exception)
            //    {                   
            //    }
            //}
            if (excelApp != null)
            {
                try
                {
                    excelApp.Quit();
                    Marshal.FinalReleaseComObject(excelApp);
                    excelApp = null;
                    GC.Collect();
                }
                catch (Exception)
                {
                }
            }

        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Ecodistrict.Excel
{
    /// <summary>
    /// The class that handles comunication with Excel. After that it waits until it is ordered to open an existing workbook.
    /// 
    /// </summary>
    /// <remarks> 
    /// Usage: <br/>
    /// 1. Create a new instance of CExcel. (CExcel cExcel=new CExcel)<br/>
    /// 2. Open a existing Excel document (.xls or .xlsx) using the complete path to it (bool res = cExcel.OpenWorkBook(path))<br/>
    /// 3. Read or Write to the document using some of the following functions: <br/>
    /// <tb/>object obj = GetCellValue("Sheet1","A1")<br/>
    /// <tb/>object obj = GetCellValue("Sheet1",row, column)<br/>
    /// <tb/>bool res = SetCellValue("Sheet1","A1", obj)<br/>
    /// <tb/>object obj = GetCellVAlue("Sheet1",,row, column, obj)<br/>
    /// 4. Close Excel and documents using CloseExcel()<br/>    
    /// </remarks> 

    public class CExcel
    {
        private string WorkBookFilePath { get; set; }
        /// <summary>
        /// The Excel application object, com object
        /// </summary>
        private MsExcel.Application _excelApp;
        /// <summary>
        /// Excel workbook object
        /// </summary>
        private MsExcel.Workbook _workbook;

        /// <summary>
        /// Excel worksheet object
        /// </summary>
        private MsExcel.Worksheet _worksheet;

        /// <summary>
        /// The constructor creates a new MsExcel.Application object
        /// </summary>
        public CExcel()
        {
            try
            {
                _excelApp = new MsExcel.Application();
            }
            catch (Exception)
            {
                
            }
        }

        /// <summary>
        /// The finalize method tries to close the Exceldocumnt and the Excel application object if this has not been done already
        /// </summary>
        ~CExcel()
        {
            CloseExcel();
        }

        /// <summary>
        /// Closes all connection to Excel and used Excel documents. 
        /// </summary>
        public void CloseExcel()
        {
            GC.Collect();
            if (_worksheet != null)
            {
                try
                {
                    _excelApp.DisplayAlerts = false;
                    Marshal.FinalReleaseComObject(_worksheet);
                    _worksheet = null;
                    GC.Collect();
                }
                catch (Exception)
                {
                }
            }
            if (_workbook != null)
            {
                try
                {
                    _excelApp.DisplayAlerts = false;
                    _workbook.Close(false, Type.Missing, Type.Missing);
                    Marshal.FinalReleaseComObject(_workbook);
                    _workbook = null;
                    GC.Collect();
                }
                catch (Exception)
                {
                }
            }
            if (_excelApp != null)
            {
                try
                {
                    _excelApp.Quit();
                    Marshal.FinalReleaseComObject(_excelApp);
                    _excelApp = null;
                    GC.Collect();
                }
                catch (Exception)
                {
                }
            }

        }
        /// <summary>
        /// Opens a Excel document file (.xls or .xlsx) 
        /// </summary>
        /// <param name="path">A complete path to an existing Excel document file. (C:\library\myExcelfile.xlsx)</param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        public bool OpenWorkBook(string path)
        {
            try 
            { 
                if (_excelApp == null)
                {
                    _excelApp=new MsExcel.Application();
                }
            }
            catch(Exception)
            {
                return false;
            }

            if (!File.Exists(path))
            {
                return false;
            }

            try
            {
               _workbook = _excelApp.Workbooks.Open(path, ReadOnly: true);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Closes an open Excel workbook without saving any changes in it.
        /// </summary>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        public bool CloseWorkBook()
        {
            try
            {
                if (_workbook != null)
                {
                    _worksheet = null;

                    _workbook.Close(SaveChanges: false);
                    _workbook = null;

                    GC.Collect();
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Sets the value using the specified parameters, using row and column syntax 
        /// </summary>
        /// <param name="sheet">Sheetname, string</param>
        /// <param name="row">Row, integer</param>
        /// <param name="col">Column, integer</param>
        /// <param name="value">Value, Object</param>
        /// <returns><see cref="Boolean">true</see> if success, <see cref="Boolean">false</see> if not</returns>
        public bool SetCellValue(string sheet, int row, int col, object value)
        {
            try
            {
                if (_worksheet == null || _worksheet.Name != sheet)
                    _worksheet = _workbook.Worksheets[sheet];

                _worksheet.Range[row, col].Value = value;
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
                if (_worksheet == null || _worksheet.Name != sheet)
                    _worksheet = _workbook.Worksheets[sheet];

                _worksheet.Range[cell].Value = value;
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
                if (_worksheet == null || _worksheet.Name != sheet)
                    _worksheet = _workbook.Worksheets[sheet];

                //celltype = ws.Range[cell].Value.GetType();
                return _worksheet.Range[cell].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} cell:{1}", sheet, cell));
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
                if (_worksheet == null || _worksheet.Name != sheet)
                    _worksheet = _workbook.Worksheets[sheet];

                //celltype = ws.Range[row, col].Value.GetType();
                return _worksheet.Range[row, col].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} row:{1} col:{2}", sheet, row, col));
            }
        }

    }
}

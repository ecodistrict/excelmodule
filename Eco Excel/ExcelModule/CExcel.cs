using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace Ecodistrict.Excel
{
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
        private MsExcel.Worksheet _worksheets;

        public CExcel()
        {
            try
            {
                _excelApp = new MsExcel.Application();
            }
            catch (Exception ex)
            {
                
            }
        }

        ~CExcel()
        {
            CloseExcel();
        }

        public void CloseExcel()
        {
            GC.Collect();
            if (_worksheets != null)
            {
                try
                {
                    _excelApp.DisplayAlerts = false;
                    Marshal.FinalReleaseComObject(_worksheets);
                    _worksheets = null;
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

        public bool OpenWorkBook(string path)
        {
            try 
            { 
                if (_excelApp == null)
                {
                    _excelApp=new MsExcel.Application();
                }
            }
            catch(Exception ex)
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
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        public bool CloseWorkBook()
        {
            try
            {
                if (_workbook != null)
                    _workbook.Close(SaveChanges: false);
            }
            catch (Exception Ex)
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
                if (_worksheets == null || _worksheets.Name != sheet)
                    _worksheets = _workbook.Worksheets[sheet];

                _worksheets.Range[row, col].Value = value;
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
                if (_worksheets == null || _worksheets.Name != sheet)
                    _worksheets = _workbook.Worksheets[sheet];

                _worksheets.Range[cell].Value = value;
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
                if (_worksheets == null || _worksheets.Name != sheet)
                    _worksheets = _workbook.Worksheets[sheet];

                //celltype = ws.Range[cell].Value.GetType();
                return _worksheets.Range[cell].Value;

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
                if (_worksheets == null || _worksheets.Name != sheet)
                    _worksheets = _workbook.Worksheets[sheet];

                //celltype = ws.Range[row, col].Value.GetType();
                return _worksheets.Range[row, col].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} row:{1} col:{2}", sheet, row, col));
            }
        }

    }
}

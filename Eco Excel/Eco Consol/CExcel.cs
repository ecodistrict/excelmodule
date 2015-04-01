using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel=Microsoft.Office.Interop.Excel;


namespace EcoExcel
{
    public class CExcel 
    {
        public string WorkBookFilePath { get; private set; }
        private Excel.Application excelApp;
        private Excel.Workbook wb;
        private Excel.Worksheet ws;

        public CExcel(string workBookFilePath)
        {
            WorkBookFilePath = workBookFilePath;
            
            try
            {
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Open(WorkBookFilePath,ReadOnly:true);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        ~CExcel()
        {
            if (excelApp != null)
            {
                excelApp.DisplayAlerts = false;
                wb.Close(SaveChanges:false);
                //mExcel.Workbooks.Close();
                excelApp.Quit();
                excelApp = null;
            }
        }

        public bool SetCellValue(string sheet, int row, int col,object value)
        {
            try
            {
                if (ws == null || ws.Name != sheet)
                    ws = wb.Worksheets[sheet];

                ws.Range[row, col].Value = value;
                return true;
            }
            catch (Exception)
            {
                
                return false;
            }
        }
        public bool SetCellValue(string sheet, string cell, object value)
        {
            try
            {
                if (ws == null || ws.Name != sheet)
                    ws = wb.Worksheets[sheet];

                ws.Range[cell].Value = value;
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }

        public object GetCellValue(string sheet, string cell)
        {
            try
            {
                if (ws == null || ws.Name != sheet)
                    ws = wb.Worksheets[sheet];

                //celltype = ws.Range[cell].Value.GetType();
                return ws.Range[cell].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} cell:{1}" , sheet, cell));
            }
        }
        public object GetCellValue(string sheet, int row, int col)
        {
            try
            {
                if (ws == null || ws.Name != sheet)
                    ws = wb.Worksheets[sheet];

                //celltype = ws.Range[row, col].Value.GetType();
                return ws.Range[row, col].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} row:{1} col:{2}", sheet, row, col));
            }
        }

        public void CloseExcel()
        {
            GC.Collect();
            if (ws != null)
            {
                excelApp.DisplayAlerts = false;
                Marshal.FinalReleaseComObject(ws);
                ws = null;
                GC.Collect();
            }
            if (wb != null)
            {
                excelApp.DisplayAlerts = false;
                wb.Close(false, Type.Missing,Type.Missing);
                Marshal.FinalReleaseComObject(wb);
                wb = null;
                GC.Collect();
            }
            if (excelApp.Workbooks != null)
            {
                excelApp.Workbooks.Close();
                Marshal.FinalReleaseComObject(excelApp.Workbooks);
                GC.Collect();
            }
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
                excelApp = null;
                GC.Collect();
            }

        }
    }
}
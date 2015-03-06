using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IMB;
using Excel=Microsoft.Office.Interop.Excel;


namespace EcoExcel
{
    public class CExcel 
    {
        public string WorkBookFilePath { get; private set; }
        private Excel.Application mExcel;
        private Excel.Workbook wb;
        private Excel.Worksheet ws;

        public CExcel(string workBookFilePath)
        {
            WorkBookFilePath = workBookFilePath;
            if (!File.Exists(WorkBookFilePath))
            {
                throw new ArgumentException("Excelfile not found");
            }
            try
            {
                mExcel = new Microsoft.Office.Interop.Excel.Application();
                wb = mExcel.Workbooks.Open(WorkBookFilePath,ReadOnly:true);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        ~CExcel()
        {
            if (mExcel != null)
            {
                mExcel.DisplayAlerts = false;
                wb.Close(SaveChanges:false);
                //mExcel.Workbooks.Close();
                mExcel.Quit();
                mExcel = null;
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

        public object GetCellValue(string sheet, string cell, out string celltype)
        {
            try
            {
                if (ws == null || ws.Name != sheet)
                    ws = wb.Worksheets[sheet];

                celltype = ws.Range[cell].Value.GetType();
                return ws.Range[cell].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} cell:{1}" , sheet, cell));
            }
        }
        public object GetCellValue(string sheet, int row, int col, out string celltype)
        {
            try
            {
                if (ws == null || ws.Name != sheet)
                    ws = wb.Worksheets[sheet];

                celltype = ws.Range[row, col].Value.GetType();
                return ws.Range[row, col].Value;

            }
            catch (Exception)
            {
                throw new Exception(string.Format("Could not read worksheet:{0} row:{1} col:{2}", sheet, row, col));
            }
        }
    }
}
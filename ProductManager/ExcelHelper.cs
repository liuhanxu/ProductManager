
using System;
using System.Data;
using System.Configuration;
using System.Web;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;


namespace ProductManager
{
    /// <SUMMARY>
    /// ExcelHelper
    /// </SUMMARY>
    public class ExcelHelper
    {
        public string mFilename;
        public Application app;
        public Workbooks wbs;
        public Workbook wb;
        public Worksheets wss;
        public Worksheet ws;
        public ExcelHelper()
        {
            app = new Application();
            wbs = app.Workbooks;
            //wb = wbs.Add(true);
        }

        public void Open(string FileName)
        {
            
            //wb = wbs.Add(FileName);
            wb = wbs.Open(FileName);
            mFilename = FileName;
        }
        public Worksheet GetSheet(string SheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets[SheetName];
            return s;
        }
        public Worksheet AddSheet(string SheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            s.Name = SheetName;
            return s;
        }

        public void DelSheet(string SheetName)
        {
            ((Worksheet)wb.Worksheets[SheetName]).Delete();
        }
        public Worksheet ReNameSheet(string OldSheetName, string NewSheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets[OldSheetName];
            s.Name = NewSheetName;
            return s;
        }

        public Worksheet ReNameSheet(Worksheet Sheet, string NewSheetName)
        {
            Sheet.Name = NewSheetName;
            return Sheet;
        }

        public void SetCellValue(Worksheet ws, int x, int y, object value)
        {
            ws.Cells[x, y] = value;
        }
        public void SetCellValue(string ws, int x, int y, object value)
        {
            GetSheet(ws).Cells[x, y] = value;
        }

      
        public bool Save()
        {
            Console.WriteLine(wb.FullName);
            if (mFilename == "")
            {
                return false;
            }
            else
            {
                try
                {
                    app.DisplayAlerts = false;
                    wb.Save();
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("save error:"+ex.Message);
                    return false;
                }
            }
        }
        public bool SaveAs(object FileName)
        {
            try
            {
                wb.Saved = true;
                wb.SaveCopyAs(FileName);
                app.DisplayAlerts = false;
                wb.SaveAs(FileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public void Close()
        {
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }
    }
}
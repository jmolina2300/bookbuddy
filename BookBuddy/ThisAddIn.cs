using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace BookBuddy
{
    public partial class ThisAddIn
    {
        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            
            //Excel.Range firstRow = activeWorksheet.get_Range("A1",missing);                  // Troublesome
            //firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, missing);    // Troublesome
            //Excel.Range newFirstRow = activeWorksheet.get_Range("A1", missing);              // Troublesome
            //newFirstRow.Value2 = "This text was added by using code";
        }
        public Excel.Worksheet GetActiveWorkSheet()
        {
            return ((Excel.Worksheet)Application.ActiveSheet);
        }
        public Excel.Workbook GetActiveWorkbook()
        {
            return (Excel.Workbook)Application.ActiveWorkbook;
        }
        public void SetActiveWorkSheet(Excel.Worksheet ws)
        {
            ws.Copy( missing, Application.ActiveSheet);
            //Application.ActiveWorkbook.Worksheets.Copy
            //newWorksheet = (Excel.Worksheet)Globals.ThisWorkbook.Worksheets.Add();
            //Application.ActiveSheet = ws;
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

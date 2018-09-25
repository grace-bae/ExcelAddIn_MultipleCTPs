using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;

namespace ExcelAddIn1
{
    /*
     * Example of Excel Add-In compatible with Excel 2013+
     * Creates custom task panes for each workbook, and hides them for inactive windows
    */
    public partial class ThisAddIn
    {
        private MyUserControl myUserControl1;
        private Dictionary<String, CustomTaskPane> CTPDictionary = new Dictionary<string, CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Event Handlers

            ((Excel.AppEvents_Event)Application).NewWorkbook += ThisAddIn_NewWorkbook; // new wb created
            //Application.WorkbookActivate += Application_WorkbookActivate; // wb is in active window...
            Application.WorkbookDeactivate += Application_WorkbookDeactivate; // wb is no longer active window
            //Application.WorkbookBeforeClose += ApplicationOnWorkbookBeforeClose; // before wb closes...
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_NewWorkbook(Excel.Workbook wb)
        {
            myUserControl1 = new MyUserControl();
            CustomTaskPane tempCTP;
            tempCTP = CustomTaskPanes.Add(myUserControl1, wb.FullName);
            tempCTP.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            tempCTP.Visible = false;
            CTPDictionary.Add(wb.FullName, tempCTP);
        }

        private void Application_WorkbookDeactivate(Excel.Workbook wb)
        {
            CustomTaskPane tempCTP;
            tempCTP = CTPDictionary[wb.FullName];
            tempCTP.Visible = false; // set it back to hidden when inactive
            CTPDictionary.Add(wb.FullName, tempCTP);
        }

        public Excel.Workbook GetActiveWorkbook()
        {
            return Application.ActiveWorkbook;
        }

        public void OnLogInButtonPress()
        {
            String key;
            CustomTaskPane visibleCTP;

            key = GetActiveWorkbook().FullName;
            visibleCTP = CTPDictionary[key]; // get CTP associated with that workbook
            visibleCTP.Visible = true; // make it visible
            CTPDictionary[key] = visibleCTP; // set it as new CTP value in dictionary
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

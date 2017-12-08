using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using NovenaLibrary;
using NovenaLibrary.Config;

namespace NovenaReportingAddIn
{
    public partial class ThisAddIn
    {
        public NovenaReportingAPI novenaReportingAPI;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            // Test if workbook has Properties sheet and if Type is "Report"
            try
            {
                Excel.Worksheet propertySheet = Wb.Worksheets["properties"];
                string referredRange = Wb.Names.Item("Type").RefersTo;
                referredRange = referredRange.Replace("=", "");
                string doodlesType = propertySheet.Range[referredRange].Value;
                if (!doodlesType.Equals("Report"))
                {
                    HideRibbon();
                }
            }
            catch (Exception)
            {
                HideRibbon();
                return;
            }

            // Check if there is an application/process for each workbook
            Excel.Workbooks books = Globals.ThisAddIn.Application.Workbooks;
            if (books.Count == 1)
            {
                ConfigureNovenaReporting();
            }
            else
            {
                try
                {
                    //If there are multiple books for this application/process, then close workbook that was just opened and then reopen it with a new process/application.
                    string filePath = Wb.FullName;
                    Wb.Application.DisplayAlerts = false;
                    Wb.Close();
                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = true;
                    excelApp.DisplayFullScreen = true;
                    excelApp.DisplayFormulaBar = true;
                    excelApp.Workbooks.Open(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                }
            }
        }

        void Application_SheetChange(object Sh, Excel.Range Target)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Sh;
            if (sheet.Name == "properties")
            {
                Application.EnableEvents = false;
                MessageBox.Show("This sheet cannot be edited", "Don't Touch This Sheet", MessageBoxButtons.OK);
                Excel._Application app = sheet.Application;
                app.Undo();
                Application.EnableEvents = true;
            }
        }

        private void ConfigureNovenaReporting()
        {
            try
            {
                var excelApp = Globals.ThisAddIn.Application;
                var connectionString = Properties.Settings.Default.ConnectionString;
                var availableTablesSQL = Properties.Settings.Default.AvailableTablesSQL;
                var databaseType = (DatabaseType)Properties.Settings.Default.DatabaseType;

                novenaReportingAPI = new NovenaReportingAPI(excelApp, connectionString, availableTablesSQL, databaseType);
            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show("A dictionary property is malformed.  Make sure the property has an even number of '|' occurrences.", "Property Malformed", MessageBoxButtons.OK);
                HideRibbon();
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ".  " + ex.InnerException.Message, "Properties Worksheet", MessageBoxButtons.OK);
                HideRibbon();
                return;
            }
        }

        private void HideRibbon()
        {
            Globals.Ribbons.NovenaReporting.tab_novenaReporting.Visible = false;
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
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            this.Application.SheetChange += new Excel.AppEvents_SheetChangeEventHandler(Application_SheetChange);
        }

        #endregion
    }
}

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
using System.IO;

namespace NovenaReportingAddIn
{
    public partial class ThisAddIn
    {
        public NovenaReportingAPI novenaReportingAPI;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ShowOnlyEditConfigurationButton();
            // load NovenaFunctions.xlam
            //var novenaFunctionsAddIn = Application.AddIns.Add("~/NovenaFunctions/NovenaFunctions.xlam", true);
            //var rootPath = Path.GetFullPath(Path.Combine(new string[] { AppDomain.CurrentDomain.BaseDi‌​rectory, "..\\..\\" }));
            //var novenaFunctionsAddIn = Application.AddIns.Add(rootPath + "NovenaFunctions\\NovenaFunctions.xlam", true);
            //var novenaFunctionsAddIn = Application.AddIns.Add("C:\\Users\\Public\\Repos\\NovenaReportingAddIn\\NovenaReportingAddIn\\NovenaFunctions\\NovenaFunctions.xlam", true);
            //var novenaFunctionsAddIn = Application.AddIns.Add("C:\\NovenaFunctions.xlam", true);
            //novenaFunctionsAddIn.Installed = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            // Test if workbook has Properties sheet and if Type is "Report"
            if (IsNovenaReportingWorkbook(Wb))
            {
                // If workbook is a Doodles Reporting workbook, show ribbon.
                ShowRibbon();
            }
            else
            {
                // If workbook is not a Doodles Reporting workboook, exit function (ribbon is already hidden).
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

        private bool IsNovenaReportingWorkbook(Excel.Workbook Wb)
        {
            try
            {
                Excel.Worksheet propertySheet = Wb.Worksheets["properties"];
                string referredRange = Wb.Names.Item("Type").RefersTo;
                referredRange = referredRange.Replace("=", "");
                string doodlesType = propertySheet.Range[referredRange].Value;

                if (doodlesType.Equals("Report"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
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
                ShowOnlyEditConfigurationButton();
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                ShowOnlyEditConfigurationButton();
                return;
            }
        }

        private void HideRibbon()
        {
            Globals.Ribbons.NovenaReporting.tab_novenaReporting.Visible = false;
        }

        private void ShowOnlyEditConfigurationButton()
        {
            Globals.Ribbons.NovenaReporting.group_authentication.Visible = false;
            Globals.Ribbons.NovenaReporting.group_cellMapping.Visible = false;
            Globals.Ribbons.NovenaReporting.group_queryTools.Visible = false;
        }

        private void ShowRibbon()
        {
            //Globals.Ribbons.NovenaReporting.tab_novenaReporting.Visible = true;
            Globals.Ribbons.NovenaReporting.group_authentication.Visible = true;
            Globals.Ribbons.NovenaReporting.group_cellMapping.Visible = true;
            Globals.Ribbons.NovenaReporting.group_queryTools.Visible = true;
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

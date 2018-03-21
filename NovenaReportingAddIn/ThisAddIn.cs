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
using QueryBuilder.Config;

namespace NovenaReportingAddIn
{
    public partial class ThisAddIn
    {
        public NovenaReportingAPI novenaReportingAPI;
        //private readonly string NOVENA_XML_NAMESPACE = "http://www.w3.org/2001/XMLSchema";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ShowOnlyEditConfigurationButton();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            // Test if workbook has Properties sheet and if Type is "Report"
            var workbookPropertiesXML = GetNovenaReportingWorkbookPropertiesXML(Wb);
            if (workbookPropertiesXML != null)
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
                ConfigureNovenaReportingAPI(workbookPropertiesXML);
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

        private string GetNovenaReportingWorkbookPropertiesXML(Excel.Workbook Wb)
        {

            Office.CustomXMLParts allXMLParts = Wb.CustomXMLParts;
            foreach (Office.CustomXMLPart part in allXMLParts)
            {
                if (part.DocumentElement.BaseName == "WorkbookProperties")
                {
                    return part.XML;
                }
            }

            return null;

        }

        private void ConfigureNovenaReportingAPI(string workbookPropertiesXML)
        {
            try
            {
                var excelApp = Globals.ThisAddIn.Application;
                var connectionString = Properties.Settings.Default.ConnectionString;
                var availableTablesSQL = Properties.Settings.Default.AvailableTablesSQL;
                var databaseType = (DatabaseType)Properties.Settings.Default.DatabaseType;

                novenaReportingAPI = new NovenaReportingAPI(excelApp, connectionString, availableTablesSQL, databaseType, workbookPropertiesXML);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "  "+ ex.InnerException, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
;        }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            var xml = "";
            //var newXml = novenaReportingAPI._workbookPropertiesConfig.SerializeXML();

            // loop thru each custom xml part and delete it if it has the WorkbookProperties baseElement.
            foreach (Office.CustomXMLPart part in Wb.CustomXMLParts)
            {
                if (part.DocumentElement.BaseName == "WorkbookProperties") part.Delete();
            }

            try
            {
                Wb.CustomXMLParts.Add(xml);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Windows.Forms;

namespace NovenaReportingAddIn
{
    public partial class NovenaReporting
    {
        private void NovenaReporting_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                var novenaFunctionsAddIn = Globals.ThisAddIn.Application.AddIns.Item["NovenaFuctions2"];
                if ( novenaFunctionsAddIn == null)
                {
                    var rootPath = Path.GetFullPath(Path.Combine(new string[] { AppDomain.CurrentDomain.BaseDi‌​rectory, "..\\..\\" }));
                    novenaFunctionsAddIn = Globals.ThisAddIn.Application.AddIns.Add(rootPath + "NovenaFunctions\\NovenaFunctions2.xlam");
                    //var novenaFunctionsAddIn = Application.AddIns.Add("C:\\NovenaFunctions.xlam");
                    novenaFunctionsAddIn.Installed = true;
                }
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                Globals.ThisAddIn.Application.AddIns.Item["NovenaFunctions2"].Installed = true;
                Globals.ThisAddIn.Application.DisplayAlerts = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error loading the NovenaFunctions Excel Add-in.  " + ex.Message, 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void button_queryCreator_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.novenaReportingAPI.ShowSqlCreator();
        }

        private void button_signIn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.novenaReportingAPI.LogIn();
        }

        private void button_editConfiguration_Click(object sender, RibbonControlEventArgs e)
        {
            var dbConnStrings = Properties.Settings.Default.ConnectionStrings.Cast<string>().ToList();
            var newSettings = Globals.ThisAddIn.novenaReportingAPI.EditConfiguration(dbConnStrings);

            // Updates application settings
            if (newSettings != null)
            {
                // Create string array and then pass to new StringCollection
                var connStringsList = (List<string>)newSettings["dbConnStrings"];
                var connStringsArray = connStringsList.ToArray();
                var newStringCollection = new StringCollection();
                newStringCollection.AddRange(connStringsArray);

                // Update application settings
                Properties.Settings.Default.ConnectionStrings = newStringCollection;
                Properties.Settings.Default.ConnectionString = newSettings["activeConnectionString"].ToString();
                Properties.Settings.Default.DatabaseType = (int)newSettings["activeDatabaseType"];
                Properties.Settings.Default.AvailableTablesSQL = newSettings["availableTablesSQL"].ToString();
                Properties.Settings.Default.Save();
                Properties.Settings.Default.Reload();
            }
        }

        private void button_refresh_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.novenaReportingAPI.RefreshData();
        }

        private void button_drilldown_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.novenaReportingAPI.Drilldown();
        }
    }
}

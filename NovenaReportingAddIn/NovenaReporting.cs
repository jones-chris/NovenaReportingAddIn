using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections;
using System.Collections.Specialized;

namespace NovenaReportingAddIn
{
    public partial class NovenaReporting
    {
        private void NovenaReporting_Load(object sender, RibbonUIEventArgs e)
        {

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
    }
}

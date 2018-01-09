using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Windows.Forms;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace NovenaReportingAddIn
{
    public partial class NovenaReporting
    {
        //public Timer timer = new Timer();

        private void NovenaReporting_Load(object sender, RibbonUIEventArgs e)
        {
            //timer.Enabled = true;
            //timer.Interval = 10000; // 10 seconds
            //timer.Tick += (anotherSender, anotherE) => LoadExcelAddIns();
        }

        //private void LoadExcelAddIns()
        //{
        //    MSExcel.AddIn novenaFunctionsAddIn = null;

        //    // determine if NovenaFunctions.xlam is already in AddIns collection
        //    try
        //    {
        //        //timer.Enabled = false;
        //        novenaFunctionsAddIn = Globals.ThisAddIn.Application.AddIns.Item["NovenaFuctions"];
        //        // If an exception is not thrown, then make sure the AddIn is installed.
        //        novenaFunctionsAddIn.Installed = true;
        //    }
        //    catch
        //    {
        //        // If exception is throw, then add NovenaFunctions Excel AddIn and install it.
        //        try
        //        {
        //            //timer.Enabled = false;
        //            var rootPath = Path.GetFullPath(Path.Combine(new string[] { AppDomain.CurrentDomain.BaseDi‌​rectory, "..\\..\\" }));
        //            novenaFunctionsAddIn = Globals.ThisAddIn.Application.AddIns.Add(rootPath + "NovenaFunctions\\NovenaFunctions.xlam", true);
        //            novenaFunctionsAddIn.Installed = true;
        //        }
        //        catch (Exception ex)
        //        {
        //            //timer.Enabled = false;
        //            MessageBox.Show("There was a problem adding or installing the NovenaFunctions Excel AddIn.  " + ex.Message,
        //                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //    }
        //}

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

        private void button_setDrilldownColumns_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.novenaReportingAPI.SetDrilldownColumns();
        }
    }
}

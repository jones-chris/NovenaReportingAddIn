using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

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
    }
}

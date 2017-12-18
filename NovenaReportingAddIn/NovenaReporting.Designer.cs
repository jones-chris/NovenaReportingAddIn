namespace NovenaReportingAddIn
{
    partial class NovenaReporting : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public NovenaReporting()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab_novenaReporting = this.Factory.CreateRibbonTab();
            this.group_authentication = this.Factory.CreateRibbonGroup();
            this.group_cellMapping = this.Factory.CreateRibbonGroup();
            this.group_queryTools = this.Factory.CreateRibbonGroup();
            this.group_reportBuilder = this.Factory.CreateRibbonGroup();
            this.button_signIn = this.Factory.CreateRibbonButton();
            this.button_addCellMapping = this.Factory.CreateRibbonButton();
            this.button_deleteCellMapping = this.Factory.CreateRibbonButton();
            this.button_queryCreator = this.Factory.CreateRibbonButton();
            this.button_refresh = this.Factory.CreateRibbonButton();
            this.button_drilldown = this.Factory.CreateRibbonButton();
            this.button_setDrilldownColumns = this.Factory.CreateRibbonButton();
            this.button_checkReport = this.Factory.CreateRibbonButton();
            this.button_editConfiguration = this.Factory.CreateRibbonButton();
            this.tab_novenaReporting.SuspendLayout();
            this.group_authentication.SuspendLayout();
            this.group_cellMapping.SuspendLayout();
            this.group_queryTools.SuspendLayout();
            this.group_reportBuilder.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_novenaReporting
            // 
            this.tab_novenaReporting.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_novenaReporting.Groups.Add(this.group_authentication);
            this.tab_novenaReporting.Groups.Add(this.group_cellMapping);
            this.tab_novenaReporting.Groups.Add(this.group_queryTools);
            this.tab_novenaReporting.Groups.Add(this.group_reportBuilder);
            this.tab_novenaReporting.Label = "Novena Reporting";
            this.tab_novenaReporting.Name = "tab_novenaReporting";
            // 
            // group_authentication
            // 
            this.group_authentication.Items.Add(this.button_signIn);
            this.group_authentication.Label = "Authentication";
            this.group_authentication.Name = "group_authentication";
            // 
            // group_cellMapping
            // 
            this.group_cellMapping.Items.Add(this.button_addCellMapping);
            this.group_cellMapping.Items.Add(this.button_deleteCellMapping);
            this.group_cellMapping.Label = "Cell Mapping";
            this.group_cellMapping.Name = "group_cellMapping";
            // 
            // group_queryTools
            // 
            this.group_queryTools.Items.Add(this.button_queryCreator);
            this.group_queryTools.Items.Add(this.button_refresh);
            this.group_queryTools.Items.Add(this.button_drilldown);
            this.group_queryTools.Items.Add(this.button_setDrilldownColumns);
            this.group_queryTools.Label = "Query Tools";
            this.group_queryTools.Name = "group_queryTools";
            // 
            // group_reportBuilder
            // 
            this.group_reportBuilder.Items.Add(this.button_checkReport);
            this.group_reportBuilder.Items.Add(this.button_editConfiguration);
            this.group_reportBuilder.Label = "Report Builder";
            this.group_reportBuilder.Name = "group_reportBuilder";
            // 
            // button_signIn
            // 
            this.button_signIn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_signIn.Label = "Sign In";
            this.button_signIn.Name = "button_signIn";
            this.button_signIn.OfficeImageId = "AccessTableContacts";
            this.button_signIn.ShowImage = true;
            this.button_signIn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_signIn_Click);
            // 
            // button_addCellMapping
            // 
            this.button_addCellMapping.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_addCellMapping.Label = "Add Mapping";
            this.button_addCellMapping.Name = "button_addCellMapping";
            this.button_addCellMapping.OfficeImageId = "FieldList";
            this.button_addCellMapping.ShowImage = true;
            // 
            // button_deleteCellMapping
            // 
            this.button_deleteCellMapping.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_deleteCellMapping.Label = "Delete Mapping";
            this.button_deleteCellMapping.Name = "button_deleteCellMapping";
            this.button_deleteCellMapping.OfficeImageId = "DatasheetNewField";
            this.button_deleteCellMapping.ShowImage = true;
            // 
            // button_queryCreator
            // 
            this.button_queryCreator.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_queryCreator.Label = "Query Creator";
            this.button_queryCreator.Name = "button_queryCreator";
            this.button_queryCreator.OfficeImageId = "ViewsAdpDiagramSqlView";
            this.button_queryCreator.ShowImage = true;
            this.button_queryCreator.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_queryCreator_Click);
            // 
            // button_refresh
            // 
            this.button_refresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_refresh.Label = "Refresh Query";
            this.button_refresh.Name = "button_refresh";
            this.button_refresh.OfficeImageId = "RecurrenceEdit";
            this.button_refresh.ShowImage = true;
            this.button_refresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_refresh_Click);
            // 
            // button_drilldown
            // 
            this.button_drilldown.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_drilldown.Label = "Drilldown";
            this.button_drilldown.Name = "button_drilldown";
            this.button_drilldown.OfficeImageId = "ZoomPrintPreviewExcel";
            this.button_drilldown.ShowImage = true;
            // 
            // button_setDrilldownColumns
            // 
            this.button_setDrilldownColumns.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_setDrilldownColumns.Label = "Set Drilldown Columns";
            this.button_setDrilldownColumns.Name = "button_setDrilldownColumns";
            this.button_setDrilldownColumns.OfficeImageId = "TableStyleRowHeaders";
            this.button_setDrilldownColumns.ShowImage = true;
            // 
            // button_checkReport
            // 
            this.button_checkReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_checkReport.Label = "Check Report";
            this.button_checkReport.Name = "button_checkReport";
            this.button_checkReport.OfficeImageId = "AcceptInvitation";
            this.button_checkReport.ShowImage = true;
            // 
            // button_editConfiguration
            // 
            this.button_editConfiguration.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_editConfiguration.Label = "Edit Configuration";
            this.button_editConfiguration.Name = "button_editConfiguration";
            this.button_editConfiguration.OfficeImageId = "FilePrepareMenu";
            this.button_editConfiguration.ShowImage = true;
            this.button_editConfiguration.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_editConfiguration_Click);
            // 
            // NovenaReporting
            // 
            this.Name = "NovenaReporting";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_novenaReporting);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.NovenaReporting_Load);
            this.tab_novenaReporting.ResumeLayout(false);
            this.tab_novenaReporting.PerformLayout();
            this.group_authentication.ResumeLayout(false);
            this.group_authentication.PerformLayout();
            this.group_cellMapping.ResumeLayout(false);
            this.group_cellMapping.PerformLayout();
            this.group_queryTools.ResumeLayout(false);
            this.group_queryTools.PerformLayout();
            this.group_reportBuilder.ResumeLayout(false);
            this.group_reportBuilder.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_novenaReporting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_authentication;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_signIn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_cellMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_addCellMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_deleteCellMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_queryTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_queryCreator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_refresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_drilldown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_setDrilldownColumns;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_reportBuilder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_checkReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_editConfiguration;
    }

    partial class ThisRibbonCollection
    {
        internal NovenaReporting NovenaReporting
        {
            get { return this.GetRibbon<NovenaReporting>(); }
        }
    }
}

using NLog;
using NLog.Config;
using NLog.Targets;

namespace OutlookExtensions
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbons.FolderContextMenu();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

            ConfigureNlog();
        }

        private void ConfigureNlog()
        {
            LoggingConfiguration config = new LoggingConfiguration();
            FileTarget fileTarget = new FileTarget("file")
            {
                FileName = "${basedir}/Logs/OutlookExtensions.log",
                Layout = "${longdate} ${level} ${message}  ${exception}",
                ArchiveEvery = FileArchivePeriod.Month,
                ArchiveFileName = "Logs/ArchiveLog.{#}.log",
                ArchiveNumbering = ArchiveNumberingMode.Date,
                ArchiveDateFormat = "yyyMMddHHmm",
                MaxArchiveFiles = 5

            };

            config.AddTarget(fileTarget);
            config.AddRuleForAllLevels("file", "*");
            LogManager.Configuration = config;
        }

        #endregion
    }
}

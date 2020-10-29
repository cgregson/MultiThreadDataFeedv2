using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Threading;

namespace MultiThreadDataFeedv2
{
    public partial class ThisAddIn
    {
        private DataFeed dataFeed;
        public string dataFeedFileName { get; private set; }
        protected override object RequestComAddInAutomationService()
        {
            if (dataFeed == null)
            {
                dataFeed = new DataFeed();
            }
            return dataFeed;
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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

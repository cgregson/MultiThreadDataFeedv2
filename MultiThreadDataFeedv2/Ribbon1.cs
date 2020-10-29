using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace MultiThreadDataFeedv2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            CreateExportThreadTest();
        }
        public void CreateExportThreadTest()
        {
            var thread = new Thread(() => ExportDataTest());
            thread.Start();
        }
        private void ExportDataTest()
        {
            var app = new Application();
            app.Visible = true;
            app.Run(@"'Y:\CityRealEstate\Acquisitions\Form Underwriting Model\VBA\LTDX - Copy.xlam'!RefreshData", Globals.ThisAddIn.Application.ActiveWorkbook);
        }
    }
}

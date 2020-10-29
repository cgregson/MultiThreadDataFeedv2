using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System;
using System.IO;

namespace MultiThreadDataFeedv2
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class DataFeed : IDataFeed
    {
        public void CreateExportThread(string workbookFullPath, bool active, bool closed, string scenario)
        {
            var thread = new Thread(() => ExportData(workbookFullPath, active, closed, scenario));
            thread.Start();
        }
        private void ExportData(string workbookFullPath, bool active, bool closed, string scenario)
        {
            var app = new Excel.Application();
            var workbook = app.Workbooks.Open(workbookFullPath, UpdateLinks: false, ReadOnly: true);
            app.Run("'" + Globals.ThisAddIn.dataFeedFileName + "'!RefreshData", workbook, null, null, null, "ExportMultiThread", null, null, null, active, closed, scenario);
            app.Workbooks[Path.GetFileName(Globals.ThisAddIn.dataFeedFileName)].Close(SaveChanges: false); //close LTDX
            workbook.Close(SaveChanges: false);
            app.Quit();
        }
        public void CreateExportThreadTest()
        {
            var thread = new Thread(() => ExportDataTest());
            thread.Start();
        }
        private void ExportDataTest()
        {
            Globals.ThisAddIn.Application.Run("'DataFeed DEV.xlam.xlsm'!ExportData");
        }
    }
}

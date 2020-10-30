using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Speech.Synthesis;
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
            Globals.ThisAddIn.Application.ActiveWorkbook.SaveCopyAs(@"C:\Users\" + Environment.UserName + @"\AppData\Local" + Globals.ThisAddIn.Application.ActiveWorkbook.Name); //save to temp folder
            var app = new Application
            {
                Visible = true
            };
            var wb = app.Workbooks.Open(@"C:\Development\" + Globals.ThisAddIn.Application.ActiveWorkbook.Name);
            var wbFullName = wb.FullName;
            try
            {
                app.Run(@"'Y:\CityRealEstate\Acquisitions\Form Underwriting Model\VBA\LTDX.xlam'!RefreshData", wb);
                using (SpeechSynthesizer synth = new SpeechSynthesizer()) //not managed by garbage collection, need using statement to dispose object
                {
                    synth.Speak(WindowsDisplayName() + ", thank you. Your data has been assimilated.");
                };
            }
            catch
            {
                using (SpeechSynthesizer synth = new SpeechSynthesizer()) //not managed by garbage collection, need using statement to dispose object
                {
                    synth.Speak(WindowsDisplayName() + ", your data export failed.");
                };
            }
            finally
            {
                wb.Close(false);
                File.Delete(wbFullName);
                app.Quit();
            }
        }
        private string WindowsDisplayName()
        {
            Thread.GetDomain().SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal);
            WindowsPrincipal principal = (WindowsPrincipal)Thread.CurrentPrincipal;
            using (PrincipalContext pc = new PrincipalContext(ContextType.Domain))
            {
                UserPrincipal up = UserPrincipal.FindByIdentity(pc, principal.Identity.Name);
                return up.DisplayName;
            }
        }
    }
}

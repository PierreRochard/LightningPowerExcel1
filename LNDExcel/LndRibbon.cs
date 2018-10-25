using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Grpc.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace LNDExcel
{
    public partial class LndRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void SetupWB_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            app.Visible = true;
            Globals.ThisAddIn.SetupWorkbook(app.ActiveWorkbook);
        }

        private void EditBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            //var address = Globals.ThisAddIn.LApp.LndClient.NewAddress();
            //editBox2.Text = address.Address;
        }
    }
}

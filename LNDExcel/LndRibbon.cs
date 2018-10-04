using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Grpc.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using TextBox = System.Windows.Forms.TextBox;

namespace LNDExcel
{
    public partial class LndRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void newAddressClick(object sender, RibbonControlEventArgs e)
        {
            var newAddress = Globals.ThisAddIn.LApp.NewAddress();
            MessageBox.Show($@"New address: {newAddress}");
        }

        private void connectLnd2_Click(object sender, RibbonControlEventArgs e)
        {
            Application app;
            try
            {
                app = (Application) Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                app = new Application();
            }
            Worksheet ws = app.ActiveSheet;
            if (ws == null)
            {
                const string message = "Open an existing LND workbook or a new workbook before connecting.";
                var result = MessageBox.Show(message);
                return;
            }

            try
            {
                Worksheet infoWorksheet = app.Sheets[SheetNames.GetInfo];
            }
            catch (COMException)
            {
                Workbook wb = app.ActiveWorkbook;
                string message = $"Initialize LNDExcel in the active workbook {wb.FullName}? This may cause data loss.";
                string caption = "";
                var result = MessageBox.Show(message, caption,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            ConnectLnd();
        }

        private void ConnectLnd()
        {
            try
            {
                Globals.ThisAddIn.ConnectLnd();
            }
            catch (RpcException rpcException)
            {
                var result = MessageBox.Show(rpcException.Status.Detail, "", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (result == DialogResult.Retry)
                {
                    ConnectLnd();
                }
            }

        }
        
    }
}

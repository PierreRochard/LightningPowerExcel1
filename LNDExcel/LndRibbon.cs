using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Google.Protobuf;
using Grpc.Core;
using Lnrpc;
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
                MessageBox.Show(message);
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

        private void editBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var address = Globals.ThisAddIn.LApp.LndClient.NewAddress();
            editBox2.Text = address.Address;
        }
    }
}

using System;
using System.Drawing;
using System.Runtime.InteropServices;

using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace LNDExcel
{
    public partial class ThisAddIn
    {

        public AsyncLightningApp LApp;
        public SendPaymentSheet SendPaymentSheet;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WorkbookOpen += ApplicationOnWorkbookOpen;
        }

        private bool IsLndWorkbook()
        {
            try
            {
                if (Application.Sheets[SheetNames.GetInfo].Cells[1, 1].Value2 == "LNDExcel")
                {
                    return true;
                }
            }
            catch (COMException)
            {
                // GetInfo tab doesn't exist, certainly not an LNDExcel workbook
            }

            return false;
        }

        // Check to see if the workbook is an LNDExcel workbook
        private void ApplicationOnWorkbookOpen(Workbook wb)
        {
            if (IsLndWorkbook())
            {
                ConnectLnd();
            }
        }

        public void ConnectLnd()
        {
            this.LApp = new AsyncLightningApp(this);
            SetupSheet(SheetNames.GetInfo);
            MarkLndExcelWorkbook();
            SetupSheet(SheetNames.Channels);
            SetupSheet(SheetNames.Payments);
            SetupSheet(SheetNames.SendPayment);
            this.SendPaymentSheet = new SendPaymentSheet(Application.Sheets[SheetNames.SendPayment], this.LApp);
            SendPaymentSheet.InitializePaymentRequest();
            LApp.Refresh(SheetNames.GetInfo);
            Application.SheetActivate += Workbook_SheetActivate;
        }

        private void SetupSheet(string worksheetName)
        {
            Worksheet oldWs = Application.ActiveSheet;
            Worksheet ws;
            try
            {
                ws = Application.Sheets[worksheetName];
            }
            catch (COMException)
            {
                Globals.ThisAddIn.Application.Sheets.Add();
                ws = Application.ActiveSheet;
                ws.Name = worksheetName;
                ws.Range["A:AZ"].Interior.Color = Color.White;
            }
            oldWs.Activate();
        }
        
        private void Workbook_SheetActivate(object sh)
        {
            if (!IsLndWorkbook())
            {
                return;
            }
            var ws = (Worksheet) sh;
            LApp.Refresh(ws.Name);
        }

        private void MarkLndExcelWorkbook()
        {
            Worksheet ws = Application.Sheets[SheetNames.GetInfo];
            ws.Cells[1, 1].Value2 = "LNDExcel";
            ws.Cells[1, 1].Font.Color = Color.White;
        }
        
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
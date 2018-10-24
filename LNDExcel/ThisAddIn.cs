﻿using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public partial class ThisAddIn
    {

        public AsyncLightningApp LApp;
        public VerticalTableSheet<GetInfoResponse> GetInfoSheet;
        public TableSheet<Channel> ChannelsSheet;
        public TableSheet<Payment> PaymentsSheet;
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
            LApp = new AsyncLightningApp(this);

            CreateSheet(SheetNames.GetInfo);
            GetInfoSheet = new VerticalTableSheet<GetInfoResponse>(Application.Sheets[SheetNames.GetInfo], LApp, GetInfoResponse.Descriptor);
            GetInfoSheet.SetupVerticalTable("LND Node Info");
            LApp.Refresh(SheetNames.GetInfo);

            CreateSheet(SheetNames.OpenChannels);
            ChannelsSheet = new TableSheet<Channel>(Application.Sheets[SheetNames.OpenChannels], LApp, Channel.Descriptor, "chan_id");
            ChannelsSheet.SetupTable("Open Channels", 3);
            LApp.Refresh(SheetNames.OpenChannels);

            CreateSheet(SheetNames.Payments);
            PaymentsSheet = new TableSheet<Payment>(Application.Sheets[SheetNames.Payments], LApp, Payment.Descriptor, "payment_hash");
            PaymentsSheet.SetupTable("Payments", 3);
            LApp.Refresh(SheetNames.Payments);

            CreateSheet(SheetNames.SendPayment);
            SendPaymentSheet = new SendPaymentSheet(Application.Sheets[SheetNames.SendPayment], LApp);
            SendPaymentSheet.InitializePaymentRequest();

            GetInfoSheet.Ws.Activate();
            MarkLndExcelWorkbook();
            
            LApp.Refresh(SheetNames.GetInfo);
            Application.SheetActivate += Workbook_SheetActivate;
        }

        private void CreateSheet(string worksheetName)
        {
            Worksheet oldWs = Application.ActiveSheet;
            Worksheet ws;
            try
            {
                // ReSharper disable once RedundantAssignment
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
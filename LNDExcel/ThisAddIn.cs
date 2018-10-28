﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Channel = Lnrpc.Channel;

namespace LNDExcel
{
    public partial class ThisAddIn
    {

        public AsyncLightningApp LApp;
        public Workbook Wb;

        public ConnectSheet ConnectSheet;
        public TableSheet<Peer> PeersSheet;
        public BalancesSheet BalancesSheet;
        public TableSheet<Channel> OpenChannelsSheet;
        public PendingChannelsSheet PendingChannelsSheet;
        public TableSheet<ChannelCloseSummary> ClosedChannelsSheet;
        public TableSheet<Payment> PaymentsSheet;
        public SendPaymentSheet SendPaymentSheet;
        public NodeSheet NodesSheet;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WorkbookOpen += ApplicationOnWorkbookOpen;
        }

        private bool IsLndWorkbook()
        {
            try
            {
                if (Application.Sheets[SheetNames.Connect].Cells[1, 1].Value2 == "LNDExcel")
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
                SetupWorkbook(wb);
            }
        }

        public void SetupWorkbook(Workbook wb)
        {
            Wb = wb;
            LApp = new AsyncLightningApp(this);

            CreateSheet(SheetNames.Connect);
            ConnectSheet = new ConnectSheet(Wb.Sheets[SheetNames.Connect], LApp);
            ConnectSheet.PopulateConfig();

            CreateSheet(SheetNames.Peers);
            PeersSheet = new TableSheet<Peer>(Wb.Sheets[SheetNames.Peers], LApp, Peer.Descriptor, "pub_key");
            PeersSheet.SetupTable("Peers", 3);

            CreateSheet(SheetNames.Balances);
            BalancesSheet = new BalancesSheet(Wb.Sheets[SheetNames.Balances], LApp);

            CreateSheet(SheetNames.OpenChannels);
            OpenChannelsSheet = new TableSheet<Channel>(Wb.Sheets[SheetNames.OpenChannels], LApp, Channel.Descriptor, "chan_id");
            OpenChannelsSheet.SetupTable("Open Channels");

            CreateSheet(SheetNames.PendingChannels);
            PendingChannelsSheet = new PendingChannelsSheet(Wb.Sheets[SheetNames.PendingChannels], LApp);

            CreateSheet(SheetNames.ClosedChannels);
            ClosedChannelsSheet = new TableSheet<ChannelCloseSummary>(Wb.Sheets[SheetNames.ClosedChannels], LApp, ChannelCloseSummary.Descriptor, "chan_id");
            ClosedChannelsSheet.SetupTable("Closed Channels");

            CreateSheet(SheetNames.Payments);
            PaymentsSheet = new TableSheet<Payment>(Wb.Sheets[SheetNames.Payments], LApp, Payment.Descriptor, "payment_hash");
            PaymentsSheet.SetupTable("Payments");

            CreateSheet(SheetNames.SendPayment);
            SendPaymentSheet = new SendPaymentSheet(Wb.Sheets[SheetNames.SendPayment], LApp);
            SendPaymentSheet.InitializePaymentRequest();

            CreateSheet(SheetNames.NodeLog);
            NodesSheet = new NodeSheet(Wb.Sheets[SheetNames.NodeLog]);

            MarkLndExcelWorkbook();
            ConnectSheet.Ws.Activate();

            Application.SheetActivate += Workbook_SheetActivate;
        }

        private void CreateSheet(string worksheetName)
        {
            Worksheet oldWs = Wb.ActiveSheet;
            Worksheet ws;
            try
            {
                // ReSharper disable once RedundantAssignment
                ws = Wb.Sheets[worksheetName];
            }
            catch (COMException)
            {
                Globals.ThisAddIn.Wb.Sheets.Add(After: Wb.Sheets[Wb.Sheets.Count]);
                ws = Wb.ActiveSheet;
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
            LApp?.Refresh(ws.Name);
        }

        private void MarkLndExcelWorkbook()
        {
            Worksheet ws = Wb.Sheets[SheetNames.Connect];
            ws.Cells[1, 1].Value2 = "LNDExcel";
            ws.Cells[1, 1].Font.Color = Color.White;
        }
        
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (!NodesSheet.isProcessOurs) return;

            try
            {
                LApp.StopDaemon();
            }
            catch (RpcException exception)
            {
                NodesSheet.isProcessOurs = false;
            }
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
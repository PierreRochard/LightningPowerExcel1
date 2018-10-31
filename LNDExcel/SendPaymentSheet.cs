using System;
using System.Collections.Generic;
using System.Drawing;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Button = Microsoft.Office.Tools.Excel.Controls.Button;

namespace LNDExcel
{
    public class SendPaymentSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        public VerticalTableSheet<PayReq> PaymentRequestTable;
        public VerticalTableSheet<Route> RouteTakenTable;
        public TableSheet<Hop> HopTable;
        public TableSheet<Route> PotentialRoutesTable;

        private const int MaxRoutes = 10;

        private Range _payReqLabelCell;
        private Range _payReqInputCell;
        private Range _payReqRange;

        private Range _errorData;

        private Range _sendStatusRange;

        private Range _paymentPreimageCell;
        private Range _paymentPreimageLabel;

        public int StartColumn = 2;
        public int StartRow = 2;

        private const int PayReqColumnWidth = 70;

        public SendPaymentSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;
        }
        
        public void InitializePaymentRequest()
        {
            _payReqLabelCell = Ws.Cells[StartRow, StartColumn];
            _payReqLabelCell.Value2 = "Payment request:";
            _payReqLabelCell.Font.Bold = true;
            _payReqLabelCell.Columns.AutoFit();

            _payReqInputCell = Ws.Cells[StartRow, StartColumn + 1];
            _payReqInputCell.Interior.Color = Color.AliceBlue;
            Formatting.WideTableColumn(_payReqInputCell);

            _payReqRange = Ws.Range[_payReqLabelCell, _payReqInputCell];
            Formatting.UnderlineBorder(_payReqRange);

            Ws.Change += WsOnChangeParsePayReq;

            PaymentRequestTable = new VerticalTableSheet<PayReq>(Ws, LApp, PayReq.Descriptor);
            PaymentRequestTable.SetupVerticalTable("Decoded Payment Request", StartRow + 2);

            PotentialRoutesTable = new TableSheet<Route>(Ws, LApp, Route.Descriptor, "hops");
            PotentialRoutesTable.SetupTable("Potential Routes", MaxRoutes, StartRow=PaymentRequestTable.EndRow + 2);
            
            var sendPaymentButtonRow = PotentialRoutesTable.EndRow + 4;
            Button sendPaymentButton = Utilities.CreateButton("sendPayment", Ws, Ws.Cells[sendPaymentButtonRow, StartColumn], "Send Payment");
            sendPaymentButton.Click += SendPaymentButtonOnClick;
            _errorData = Ws.Cells[sendPaymentButtonRow + 3, StartColumn + 1];

            _sendStatusRange = Ws.Cells[sendPaymentButtonRow + 3, StartColumn];
            _sendStatusRange.Font.Italic = true;

            Button clearPaymentInfoButton = Utilities.CreateButton("clearPaymentInfo", Ws, Ws.Cells[sendPaymentButtonRow + 6, StartColumn], "Clear");
            clearPaymentInfoButton.Click += ClearPaymentInfoButtonOnClick;
            
            var paymentResponseDataStartRow = sendPaymentButtonRow + 9;
            _paymentPreimageLabel = Ws.Cells[paymentResponseDataStartRow, StartColumn + 1];
            _paymentPreimageLabel.Value2 = "Proof of Payment";
            _paymentPreimageLabel.Font.Italic = true;

            _paymentPreimageCell = Ws.Cells[paymentResponseDataStartRow, StartColumn + 2];
            _paymentPreimageCell.Interior.Color = Color.PaleGreen;
            
            RouteTakenTable = new VerticalTableSheet<Route>(Ws, LApp, Route.Descriptor, new List<string> { "hops" });
            RouteTakenTable.SetupVerticalTable("Payment Summary", paymentResponseDataStartRow + 3);

            HopTable = new TableSheet<Hop>(Ws, LApp, Hop.Descriptor, "chan_id");
            HopTable.SetupTable("Route", 4, RouteTakenTable.EndRow + 2);

            _payReqInputCell.Columns.ColumnWidth = PayReqColumnWidth;
            Utilities.RemoveLoadingMark(Ws);
        }


        private void ClearPaymentInfoButtonOnClick(object sender, EventArgs e)
        {
            ClearPayReq();
            PaymentRequestTable.Clear();
            PotentialRoutesTable.Clear();
            ClearSendStatus();
            ClearErrorData();
            ClearSendPaymentResponseData();
        }

        private void ClearPayReq()
        {
            _payReqInputCell.Value2 = "";
        }

        private void ClearErrorData()
        {
            _errorData.Value2 = "";
            Formatting.DeactivateErrorCell(_errorData);
        }

        private void ClearSendStatus()
        {
            _sendStatusRange.Value2 = "";
        }

        private void ClearSendPaymentResponseData()
        {
            _paymentPreimageCell.Value2 = "";
            RouteTakenTable.Clear();
            HopTable.Clear();
        }

        private void DisplayError(string errorType, string errorMessage)
        {
            _errorData.Value2 = $"{errorType}: {errorMessage}";
            Formatting.ActivateErrorCell(_errorData);
        }

        private void WsOnChangeParsePayReq(Range target)
        {
            
            if (target.Address != "$C$2")
            {
                return;
            }

            string payReq = target.Value2;
            if (string.IsNullOrWhiteSpace(payReq))
            {
                return;
            }

            PayReq response;
            try
            {
                response = LApp.DecodePaymentRequest(payReq);
            }
            catch (RpcException e)
            {
                DisplayError("Parsing error", e.Status.Detail);
                return;
            }
            PaymentRequestTable.Update(response);
            ClearErrorData();

            try
            {
               var r = LApp.QueryRoutes(response, MaxRoutes);
               PotentialRoutesTable.Update(r.Routes);
            }
            catch (RpcException e)
            {
                DisplayError("Query route error", e.Status.Detail);
                return;
            }

            _payReqInputCell.Columns.ColumnWidth = PayReqColumnWidth;
        }

        private void SendPaymentButtonOnClick(object sender, EventArgs e)
        { 
            // Disable the Send Payment button so that it's not clicked twice
            Utilities.EnableButton(Ws, "sendPayment", false);

            string payReq = _payReqInputCell.Value2;
            if (string.IsNullOrWhiteSpace(payReq))
            {
                return;
            }

            try
            {
                if (PotentialRoutesTable.DataList == null || PotentialRoutesTable.DataList.Count == 0)
                {
                    LApp.SendPayment(payReq);
                }
                else
                {
                    LApp.SendPayment(payReq, PotentialRoutesTable.DataList);
                }
            }
            catch (RpcException rpcException)
            {
                DisplayError("Payment error", rpcException.Status.Detail);
            }
        }

        public void MarkSendingPayment()
        {
            ClearSendPaymentResponseData();
            ClearErrorData();
            // Indicate payment is being sent below send button
            _sendStatusRange.Value2 = "Sending payment...";
        }

        public void PopulateSendPaymentError(RpcException exception)
        {
            _errorData.Value2 = exception.Status.Detail;
            _sendStatusRange.Value2 = "";
        }

        public void PopulateSendPaymentResponse(SendResponse response)
        {
            ClearSendStatus();
            if (response.PaymentError == "")
            {
                _paymentPreimageCell.Value2 = BitConverter.ToString(response.PaymentPreimage.ToByteArray()).Replace("-", "").ToLower();

                RouteTakenTable.Populate(response.PaymentRoute);
                HopTable.Update(response.PaymentRoute.Hops);
                _payReqInputCell.Columns.ColumnWidth = PayReqColumnWidth;
            }
            else
            {
                DisplayError("Payment error", response.PaymentError);
            }
            Utilities.EnableButton(Ws, "sendPayment", true);
        }

        public void UpdateSendPaymentProgress(int progress)
        {
            // Indicate payment is being sent below send button
            _sendStatusRange.Value2 = $"Sending payment...{progress}%";
        }


    }
}
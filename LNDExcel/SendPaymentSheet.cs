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

        private Range _payReqLabelCell;
        private Range _payReqInputCell;
        private Range _payReqInputRange;
        private Range _payReqRange;

        private Range _errorData;

        private Range _sendStatusRange;

        private Range _paymentPreimageCell;
        private Range _paymentPreimageLabel;

        private int _startColumn = 2;
        private int _startRow = 2;

        private int _payReqDataStartRow = 4;
        private int _sendPaymentButtonRow = 23;
        private int _clearPaymentInfoButtonRow = 25;
        private int _paymentResponseDataStartRow = 27;

        private int _payReqColumnWidth = 70;

        public SendPaymentSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;
        }
        
        public void InitializePaymentRequest()
        {
            _payReqLabelCell = Ws.Cells[_startRow, _startColumn];
            _payReqLabelCell.Value2 = "Payment request:";
            _payReqLabelCell.Font.Bold = true;
            _payReqLabelCell.Columns.AutoFit();

            _payReqInputCell = Ws.Cells[_startRow, _startColumn + 1];

            _payReqInputRange = Ws.Range[_payReqInputCell, "U2"];
            _payReqInputRange.Interior.Color = Color.AliceBlue;

            _payReqRange = Ws.Range[_payReqLabelCell, _payReqInputRange];
            Formatting.UnderlineBorder(_payReqRange);

            Ws.Change += WsOnChangeParsePayReq;

            PaymentRequestTable = new VerticalTableSheet<PayReq>(Ws, LApp, PayReq.Descriptor);
            PaymentRequestTable.SetupVerticalTable("Decoded Payment Request", _payReqDataStartRow);

            PotentialRoutesTable = new TableSheet<Route>(Ws, LApp, Route.Descriptor, "hops");
            PotentialRoutesTable.SetupTable("Potential Routes", 3, _startRow=PaymentRequestTable.EndRow + 2);

            _errorData = Ws.Cells[_sendPaymentButtonRow + 1, _startColumn + 1];

            Button sendPaymentButton = Utilities.CreateButton("sendPayment", Ws, Ws.Cells[_sendPaymentButtonRow, _startColumn], "Send Payment");
            sendPaymentButton.Click += SendPaymentButtonOnClick;
            
            Button clearPaymentInfoButton = Utilities.CreateButton("clearPaymentInfo", Ws, Ws.Cells[_clearPaymentInfoButtonRow, _startColumn], "Clear");
            clearPaymentInfoButton.Click += ClearPaymentInfoButtonOnClick;

            _sendStatusRange = Ws.Cells[_sendPaymentButtonRow + 1, _startColumn];
            _sendStatusRange.Font.Italic = true;

            _paymentPreimageLabel = Ws.Cells[_paymentResponseDataStartRow, _startColumn + 1];
            _paymentPreimageLabel.Value2 = "Proof of Payment";
            _paymentPreimageLabel.Font.Italic = true;

            _paymentPreimageCell = Ws.Cells[_paymentResponseDataStartRow + 1, _startColumn + 1];
            _paymentPreimageCell.Interior.Color = Color.PaleGreen;
            
            RouteTakenTable = new VerticalTableSheet<Route>(Ws, LApp, Route.Descriptor, new List<string> { "hops" });
            RouteTakenTable.SetupVerticalTable("Payment Summary", _paymentResponseDataStartRow + 3);

            HopTable = new TableSheet<Hop>(Ws, LApp, Hop.Descriptor, "chan_id");
            HopTable.SetupTable("Route", 4, _paymentResponseDataStartRow + 12);

            _payReqInputCell.Columns.ColumnWidth = _payReqColumnWidth;
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
               var r = LApp.QueryRoutes(response);
               PotentialRoutesTable.Update(r.Routes);
            }
            catch (RpcException e)
            {
                DisplayError("Query route error", e.Status.Detail);
                return;
            }

            _payReqInputCell.Columns.ColumnWidth = _payReqColumnWidth;
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
                    LApp.SendPayment(PaymentRequestTable.Data);
                }
                else
                {
                    LApp.SendPayment(PaymentRequestTable.Data, PotentialRoutesTable.DataList);
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
                _payReqInputCell.Columns.ColumnWidth = _payReqColumnWidth;
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
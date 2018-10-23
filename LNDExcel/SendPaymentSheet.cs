using System;
using System.Drawing;
using System.Threading;
using Microsoft.Office.Interop.Excel;

using Grpc.Core;

using Lnrpc;

namespace LNDExcel
{
    public class SendPaymentSheet
    {


        public AsyncLightningApp LApp;
        public Worksheet Ws;

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
        private int _sendPaymentButtonRow = 16;
        private int _clearPaymentInfoButtonRow = 18;
        private int _paymentResponseDataStartRow = 20;

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

            Tables.SetupVerticalTable(Ws, "Decoded Payment Request", PayReq.Descriptor, null, _payReqDataStartRow);

            _errorData = Ws.Cells[_sendPaymentButtonRow + 1, _startColumn + 1];

            Microsoft.Office.Tools.Excel.Controls.Button sendPaymentButton = Utilities.CreateButton("sendPayment", Ws, Ws.Cells[_sendPaymentButtonRow, _startColumn], "Send Payment");
            sendPaymentButton.Click += SendPaymentButtonOnClick;
            
            Microsoft.Office.Tools.Excel.Controls.Button clearPaymentInfoButton = Utilities.CreateButton("clearPaymentInfo", Ws, Ws.Cells[_clearPaymentInfoButtonRow, _startColumn], "Clear");
            clearPaymentInfoButton.Click += ClearPaymentInfoButtonOnClick;

            _sendStatusRange = Ws.Cells[_sendPaymentButtonRow + 1, _startColumn];
            _sendStatusRange.Font.Italic = true;

            _paymentPreimageLabel = Ws.Cells[_paymentResponseDataStartRow, _startColumn + 1];
            _paymentPreimageLabel.Value2 = "Proof of Payment";
            _paymentPreimageLabel.Font.Italic = true;

            _paymentPreimageCell = Ws.Cells[_paymentResponseDataStartRow + 1, _startColumn + 1];
            _paymentPreimageCell.Interior.Color = Color.PaleGreen;
            
            Tables.SetupVerticalTable(Ws, "Payment Summary", Route.Descriptor, null, _paymentResponseDataStartRow + 3);
            Tables.SetupTable<Hop>(Ws, "Route", Hop.Descriptor, null, _paymentResponseDataStartRow + 12);

            _payReqInputCell.Columns.ColumnWidth = _payReqColumnWidth;
            Tables.RemoveLoadingMark(Ws);
        }


        private void ClearPaymentInfoButtonOnClick(object sender, EventArgs e)
        {
            ClearPayReq();
            ClearParsedPayReq();
            ClearSendStatus();
            ClearErrorData();
            ClearSendPaymentResponseData();
        }

        private void ClearPayReq()
        {
            _payReqInputCell.Value2 = "";
        }

        private void ClearParsedPayReq()
        {
            Tables.ClearVerticalTable(Ws, PayReq.Descriptor, _payReqDataStartRow);
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
            Tables.ClearVerticalTable(Ws, Route.Descriptor, _paymentResponseDataStartRow + 3);
            Tables.ClearTable(Ws, Hop.Descriptor, _paymentResponseDataStartRow + 12);
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
                response = LApp.LndClient.DecodePaymentRequest(payReq);
            }
            catch (RpcException rpcException)
            {
                DisplayError("Parsing error", rpcException.Status.Detail);
                return;
            }
            Tables.PopulateVerticalTable(Ws, PayReq.Descriptor, response, _payReqDataStartRow);

            ClearErrorData();

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
                LApp.SendPayment(payReq);
            }
            catch (RpcException rpcException)
            {
                DisplayError("Payment error", rpcException.Status.Detail);
                return;
            }
        }

        public void MarkSendingPayment()
        {
            ClearSendPaymentResponseData();
            ClearErrorData();
            // Indicate payment is being sent below send button
            _sendStatusRange.Value2 = $"Sending payment...";
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

                Tables.PopulateVerticalTable(Ws, Route.Descriptor, response.PaymentRoute, _paymentResponseDataStartRow + 3);
                Tables.PopulateTable(Ws, Hop.Descriptor, response.PaymentRoute.Hops, _paymentResponseDataStartRow + 12);
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
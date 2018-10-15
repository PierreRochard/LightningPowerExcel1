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

        private Range _errorDataLabel;
        private Range _errorData;

        private Range _sendStatusRange;

        private Range _paymentPreimageCell;
        private Range _paymentPreimageLabel;

        private int _startColumn = 2;
        private int _startRow = 2;

        private int _payReqDataStartRow = 4;
        private int _sendPaymentButtonRow = 16;
        private int _paymentResponseDataStartRow = 20;

        private int _payReqColumnWidth = 70;

        public SendPaymentSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            this.Ws = ws;
            this.LApp = lApp;
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

            Tables.PopulateVerticalTable(Ws, "Decoded Payment Request", PayReq.Descriptor, null, _payReqDataStartRow);

            _errorDataLabel = Ws.Cells[_paymentResponseDataStartRow, _startColumn];
            _errorData = Ws.Cells[_paymentResponseDataStartRow + 1, _startColumn];

            Microsoft.Office.Tools.Excel.Controls.Button sendButton = Utilities.CreateButton("sendPayment", Ws, Ws.Cells[_sendPaymentButtonRow, 2]);
            sendButton.Click += SendPaymentButtonOnClick;

            _sendStatusRange = Ws.Cells[_sendPaymentButtonRow + 1, _startColumn];
            _sendStatusRange.Font.Italic = true;

            _paymentPreimageLabel = Ws.Cells[_paymentResponseDataStartRow, _startColumn + 1];
            _paymentPreimageCell = Ws.Cells[_paymentResponseDataStartRow + 1, _startColumn + 1];

            Tables.PopulateVerticalTable(Ws, "Payment Summary", Route.Descriptor, null, _paymentResponseDataStartRow + 3);
            Tables.PopulateTable<Hop>(Ws, "Route", Hop.Descriptor, null, _paymentResponseDataStartRow + 12);

            _payReqInputCell.Columns.ColumnWidth = _payReqColumnWidth;
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
                _errorDataLabel.Value2 = "Parsing error:";
                _errorData.Value2 = rpcException.Status.Detail;
                return;
            }
            Tables.PopulateVerticalTable(Ws, "Decoded Payment Request", PayReq.Descriptor, response, _payReqDataStartRow);

            _errorData.Value2 = "";
            _errorDataLabel.Value2 = "";

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
                _errorDataLabel.Value2 = "Payment error:";
                _errorData.Value2 = rpcException.Status.Detail;
                return;
            }
        }

        public void MarkSendingPayment()
        {

            // Indicate payment is being sent below send button
            _sendStatusRange.Value2 = $"Sending payment...";

            // Clear payment response
            _paymentPreimageCell.Value2 = "";
            Tables.PopulateVerticalTable(Ws, "Payment Summary", Route.Descriptor, null, _paymentResponseDataStartRow + 3);
            Tables.PopulateTable<Hop>(Ws, "Route", Hop.Descriptor, null, _paymentResponseDataStartRow + 12);
            _payReqInputCell.Columns.ColumnWidth = _payReqColumnWidth;

            _errorData.Value2 = "";
            _errorDataLabel.Value2 = "";

        }

        public void PopulateSendPaymentError(RpcException exception)
        {
            _errorData.Value2 = exception.Status.Detail;
            _sendStatusRange.Value2 = "";
        }

        public void PopulateSendPaymentResponse(SendResponse response)
        {
            if (response.PaymentError == "")
            {
                _paymentPreimageLabel.Value2 = "Proof of Payment";
                _paymentPreimageLabel.Font.Italic = true;
                _paymentPreimageCell.Value2 = BitConverter.ToString(response.PaymentPreimage.ToByteArray()).Replace("-", "").ToLower();
                _paymentPreimageCell.Interior.Color = Color.PaleGreen;

                Tables.PopulateVerticalTable(Ws, "Payment Summary", Route.Descriptor, response.PaymentRoute, _paymentResponseDataStartRow + 3);
                Tables.PopulateTable(Ws, "Route", Hop.Descriptor, response.PaymentRoute.Hops, _paymentResponseDataStartRow + 12);
            }
            else
            {
                _errorDataLabel.Value2 = "Payment error:";
                _errorData.Value2 = response.PaymentError;
            }

            Utilities.EnableButton(Ws, "sendPayment", true);

            _sendStatusRange.Value2 = "";
            _payReqInputCell.Columns.ColumnWidth = _payReqColumnWidth;
        }

        public void UpdateSendPaymentProgress(int progress)
        {
            // Indicate payment is being sent below send button
            _sendStatusRange.Value2 = $"Sending payment...{progress}%";
        }
    }
}
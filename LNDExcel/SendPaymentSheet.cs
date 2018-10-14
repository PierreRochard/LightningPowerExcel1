using System;
using System.Drawing;

using Microsoft.Office.Interop.Excel;

using Grpc.Core;

using Lnrpc;

namespace LNDExcel
{
    public class SendPaymentSheet
    {


        public LightningApp LApp;
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

        private int _responseDataStartColumn = 2;

        private int _payReqDataStartRow = 4;
        private int _sendPaymentButtonRow = 16;
        private int _paymentResponseDataStartRow = 20;

        public SendPaymentSheet(Worksheet ws, LightningApp lApp)
        {
            this.Ws = ws;
            this.LApp = lApp;
        }
        
        public void InitializePaymentRequest()
        {
            _payReqLabelCell = Ws.Cells[2, 2];
            Ws.Names.Add("payReqLabelCell", _payReqLabelCell);
            _payReqLabelCell.Value2 = "Payment request:";
            _payReqLabelCell.Font.Bold = true;
            _payReqLabelCell.Columns.AutoFit();

            _payReqInputCell = Ws.Range["C2"];

            _payReqInputRange = Ws.Range["C2:U2"];
            _payReqInputRange.Interior.Color = Color.AliceBlue;

            _payReqRange = Ws.Range[_payReqLabelCell, _payReqInputRange];
            _payReqRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            _payReqRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

            Ws.Change += WsOnChangeParsePayReq;

            Tables.PopulateVerticalTable(Ws, "Decoded Payment Request", PayReq.Descriptor, null, _payReqDataStartRow);

            _errorDataLabel = Ws.Cells[_paymentResponseDataStartRow, _responseDataStartColumn];
            _errorData = Ws.Cells[_paymentResponseDataStartRow + 1, _responseDataStartColumn];

            Microsoft.Office.Tools.Excel.Controls.Button sendButton = Utilities.CreateButton("sendPayment", Ws, Ws.Cells[_sendPaymentButtonRow, 2]);
            sendButton.Click += SendPaymentButtonOnClick;
            _payReqInputCell.Columns.ColumnWidth = 70;

            _sendStatusRange = Ws.Cells[_sendPaymentButtonRow + 1, 2];

            _paymentPreimageLabel = Ws.Cells[_paymentResponseDataStartRow, _responseDataStartColumn + 1];
            _paymentPreimageCell = Ws.Cells[_paymentResponseDataStartRow + 1, _responseDataStartColumn + 1];
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

            target.Columns.ColumnWidth = 70;
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
            _sendStatusRange.Font.Italic = true;
            _sendStatusRange.Value2 = "Sending payment...";

            // Clear payment response
            _paymentPreimageCell.Value2 = "";
            Tables.PopulateVerticalTable(Ws, "Payment Summary", Route.Descriptor, null, _paymentResponseDataStartRow + 3);
            Tables.PopulateTable<Hop>(Ws, "Route", Hop.Descriptor, null, _paymentResponseDataStartRow + 12);
            _payReqInputCell.Columns.ColumnWidth = 70;
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
            Ws.Range["C1"].Columns.ColumnWidth = 70;
        }
    }
}
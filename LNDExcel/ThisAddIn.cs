using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace LNDExcel
{
    public interface IThisAddIn
    {
        Worksheet SetupSheet(string worksheetName);
        void MarkAsLoadingVerticalTable(string worksheetName, MessageDescriptor messageDescriptor);
        void PopulateVerticalTable(string tableTitle, string worksheetName, MessageDescriptor messageDescriptor, IMessage message, int startRow, int endRow);
    }

    public partial class ThisAddIn: IThisAddIn
    {

        public LightningApp LApp;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.WorkbookOpen += ApplicationOnWorkbookOpen;
        }

        public bool IsLndWorkbook()
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
            this.LApp = new LightningApp(this);
            SetupSheet(SheetNames.GetInfo);
            MarkLndExcelWorkbook();
            SetupSheet(SheetNames.Channels);
            SetupSheet(SheetNames.Payments);
            SetupSheet(SheetNames.SendPayment);
            InitializePaymentRequest();
            LApp.Refresh(SheetNames.GetInfo);
            Application.SheetActivate += Workbook_SheetActivate;
        }

        public Worksheet SetupSheet(string worksheetName)
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
            return ws;
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

        public void MarkLndExcelWorkbook()
        {
            Worksheet ws = Application.Sheets[SheetNames.GetInfo];
            ws.Cells[1, 1].Value2 = "LNDExcel";
            ws.Cells[1, 1].Font.Color = Color.White;
        }

        public void MarkAsLoadingTable(string worksheetName, MessageDescriptor messageDescriptor)
        {
            var ws = Application.Sheets[worksheetName];
            var fieldCount = messageDescriptor.Fields.InDeclarationOrder().Count;
            var dataRange = ws.Range[ws.Cells[2, 2], ws.Cells[100, fieldCount]];
            dataRange.Clear();
            ws.Cells[2, 2].Value2 = "Loading...";
            dataRange.Interior.Color = Color.LightGray;
        }

        public void PopulateTable<T>(string tableTitle, string worksheetName, MessageDescriptor messageDescriptor, RepeatedField<T> responseData, int startRow = 2, int startColumn = 2)
        {
            Worksheet ws = Application.Sheets[worksheetName];
            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();
            
            var endCol = fields.Count + 1;

            Range title = ws.Cells[startRow, startColumn];
            title.Font.Italic = true;
            title.Value2 = tableTitle;

            startRow++;

            var dataRange = ws.Range[ws.Cells[startRow, startColumn], ws.Cells[100, endCol]];
            dataRange.Interior.Color = Color.White;

            Range header = ws.Range[ws.Cells[startRow, startColumn], ws.Cells[startRow, endCol]];
            header.Interior.Color = Color.White;
            header.Font.Bold = true;
            header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
            header.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;

            for (var colJ = 0; colJ < fields.Count; colJ++)
            {
                var colNumber = colJ + 2;
                var headerCell = ws.Cells[startRow, colNumber];
                var field = fields[colJ];
                var fieldName = field.Name.Replace("_", " ");
                fieldName = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(fieldName);
                headerCell.Value2 = fieldName;
                headerCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                headerCell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                headerCell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                headerCell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
                headerCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                if (field.IsRepeated && field.FieldType != FieldType.Message)
                {
                    ws.Columns[colNumber].ColumnWidth = 100;
                }
            }

            for (var rowI = 0; rowI < responseData.Count; rowI++)
            {
                T data = responseData[rowI];
                var rowNumber = rowI + startRow + 1;
                for (var colJ = 0; colJ < fields.Count; colJ++)
                {
                    var field = fields[colJ];
                    var colNumber = colJ + 2;
                    var dataCell = ws.Cells[rowNumber, colNumber];

                    string value = "";
                    if (field.IsRepeated && field.FieldType != FieldType.Message)
                    {
                        var items = (RepeatedField<string>) fields[colJ].Accessor.GetValue(data as IMessage);
                        for (int i = 0; i < items.Count; i++)
                        {
                            value += items[i].ToString();
                            if (i < items.Count - 1)
                            {
                                value += ",\n";
                            }
                        }
                    }
                    else
                    {
                        value = fields[colJ].Accessor.GetValue(data as IMessage).ToString();
                    }

                    dataCell.Value2 = value;
                    dataCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    dataCell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                    dataCell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    dataCell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
                    dataCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    dataCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
                }
                Range rowRange = ws.Range[ws.Cells[rowNumber, startColumn], ws.Cells[rowNumber, endCol]];
                rowRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                rowRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                rowRange.Interior.Color = rowI % 2 == 0 ? Color.LightYellow : Color.White;
            }

            ws.Range["A:AZ"].Columns.AutoFit();
            ws.Range["A:AZ"].Rows.AutoFit();
        }

        public void InitializePaymentRequest()
        {
            Worksheet ws = Application.Sheets[SheetNames.SendPayment];
            Range payReqLabelCell = ws.Cells[2, 2];
            ws.Names.Add("payReqLabelCell", payReqLabelCell);
            payReqLabelCell.Value2 = "Payment request:";
            payReqLabelCell.Font.Bold = true;
            payReqLabelCell.Columns.AutoFit();
            Range payReqInputRange = ws.Range["C2:Q2"];
            Range payReqInputCell = ws.Range["C2"];

            ws.Names.Add("payReqInputCell", payReqInputCell);
            Range cellGroup = ws.Range[payReqLabelCell, payReqInputRange];
            cellGroup.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            cellGroup.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            payReqInputRange.Interior.Color = Color.AliceBlue;
            ws.Change += WsOnChangeParsePayReq;

            int payReqDataStartRow = 4;
            PopulateVerticalTable("Decoded Payment Request", SheetNames.SendPayment, PayReq.Descriptor, null, payReqDataStartRow);

            var buttonRow = 16;
            var buttonName = "sendPayment";
            var button = new Microsoft.Office.Tools.Excel.Controls.Button();
            var worksheet = Globals.Factory.GetVstoObject(ws);
            Range selection = ws.Cells[buttonRow, 2];
            worksheet.Controls.AddControl(button, selection, buttonName);
            button.Text = @"Send Payment";
            button.Click += SendPaymentButtonOnClick;
            payReqInputCell.Columns.ColumnWidth = 70;
        }

        private void WsOnChangeParsePayReq(Range target)
        {
            Worksheet ws = Application.Sheets[SheetNames.SendPayment];
            int payReqDataStartRow = 4;
            int responseDataStartColumn = 2;
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
                Range errorDataLabel = ws.Cells[payReqDataStartRow, responseDataStartColumn];
                Range errorData = ws.Cells[payReqDataStartRow, responseDataStartColumn + 1];
                errorDataLabel.Value2 = "Parsing error:";
                errorData.Value2 = rpcException.Status.Detail;
                return;
            }
            PopulateVerticalTable("Decoded Payment Request", SheetNames.SendPayment, PayReq.Descriptor, response, payReqDataStartRow);
            
            target.Columns.ColumnWidth = 70;
        }

        private void EnableButton(Worksheet ws, string buttonName, bool enable)
        {
            var worksheet = Globals.Factory.GetVstoObject(ws);
            foreach (Control control in worksheet.Controls)
            {
                if (control.Name == buttonName)
                {
                    control.Enabled = enable;
                }
            }
        }

        private void SendPaymentButtonOnClick(object sender, EventArgs e)
        {
            Worksheet ws = Application.Sheets[SheetNames.SendPayment];

            // Disable the Send Payment button so that it's not clicked twice
            EnableButton(ws, "sendPayment", false);

            Range inputCell = ws.Cells[2, 3];
            int paymentResponseDataStartRow = 20;
            int responseDataStartColumn = 2;

            string payReq = inputCell.Value2;
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
                Range errorDataLabel = ws.Cells[paymentResponseDataStartRow, responseDataStartColumn];
                Range errorData = ws.Cells[paymentResponseDataStartRow + 1, responseDataStartColumn];
                errorDataLabel.Value2 = "Payment error:";
                errorData.Value2 = rpcException.Status.Detail;
                return;
            }
        }

        public void MarkSendingPayment()
        {
            // Indicate payment is being sent below send button
            Worksheet ws = Application.Sheets[SheetNames.SendPayment];
            var buttonRow = 16;
            Range subtextRange = ws.Cells[buttonRow + 1, 2];
            subtextRange.Font.Italic = true;
            subtextRange.Value2 = "Sending payment...";

            // Clear payment response
               //todo
        }

        public void PopulateSendPaymentResponse(SendResponse response)
        {
            Worksheet ws = Application.Sheets[SheetNames.SendPayment];
            int responseDataStartRow = 19;
            int responseDataStartColumn = 2;

            if (response.PaymentError == "")
            {
                Range paymentPreimageLabel = ws.Cells[responseDataStartRow, responseDataStartColumn+1];
                paymentPreimageLabel.Value2 = "Proof of Payment";
                paymentPreimageLabel.Font.Italic = true;
                Range paymentPreimage = ws.Cells[responseDataStartRow+1, responseDataStartColumn+1];
                paymentPreimage.Value2 = BitConverter.ToString(response.PaymentPreimage.ToByteArray()).Replace("-", "").ToLower();
                paymentPreimage.Interior.Color = Color.PaleGreen;

                PopulateVerticalTable("Payment Summary", SheetNames.SendPayment, Route.Descriptor, response.PaymentRoute, responseDataStartRow+3);
                PopulateTable("Route", SheetNames.SendPayment, Hop.Descriptor, response.PaymentRoute.Hops, responseDataStartRow+12);
            }
            else
            {
                Range errorDataLabel = ws.Cells[responseDataStartRow, responseDataStartColumn];
                Range errorData = ws.Cells[responseDataStartRow + 1, responseDataStartColumn];
                errorDataLabel.Value2 = "Payment error:";
                errorData.Value2 = response.PaymentError;
            }
            
            EnableButton(ws, "sendPayment", true);

            var buttonRow = 16;
            Range subtextRange = ws.Cells[buttonRow + 1, 2];
            subtextRange.Value2 = "";
            ws.Range["C1"].Columns.ColumnWidth = 70;
        }

        public void MarkAsLoadingVerticalTable(string worksheetName, MessageDescriptor messageDescriptor)
        {
            Worksheet ws = Application.Sheets[worksheetName];
            ws.Select();
            var fieldCount = messageDescriptor.Fields.InDeclarationOrder().Count;
            var dataRange = ws.Range[ws.Cells[2, 2], ws.Cells[fieldCount + 2, 3]];
            dataRange.Clear();
            ws.Cells[2, 2].Value2 = "Loading...";
            dataRange.Interior.Color = Color.LightGray;
            dataRange.Columns.AutoFit();
        }

        public void PopulateVerticalTable(string tableTitle, string worksheetName, MessageDescriptor messageDescriptor, IMessage message=null, int startRow=2, int startColumn=2)
        {
            int endColumn = startColumn + 1;
            Worksheet ws = Application.Sheets[worksheetName];

            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();

            Range title = ws.Cells[startRow, startColumn];
            title.Font.Italic = true;
            title.Value2 = tableTitle;

            int dataStartRow = startRow + 1;
            int dataRow = dataStartRow;
            foreach (var field in fields)
            {
                if (field.IsRepeated)
                {
                    continue;
                }
                var fieldName = field.Name.Replace("_", " ");
                fieldName = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(fieldName);
                
                Range fieldNameCell = ws.Cells[dataRow, startColumn];
                fieldNameCell.Font.Bold = true;
                fieldNameCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                fieldNameCell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                
                Range fieldValueCell = ws.Cells[dataRow, endColumn];
                fieldNameCell.Value2 = fieldName;
                fieldValueCell.Value2 = message != null ? field.Accessor.GetValue(message).ToString() : "";
                fieldValueCell.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                Range row = ws.Range[fieldNameCell, fieldValueCell];
                row.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                row.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
                row.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                row.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                row.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                row.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
                row.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                row.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                row.Interior.Color = dataRow % 2 == 0 ? Color.LightYellow : Color.White;
                dataRow++;
            }

            ws.Range["A:D"].Columns.AutoFit();
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
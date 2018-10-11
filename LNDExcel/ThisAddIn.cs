using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
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
        void PopulateVerticalTable(string worksheetName, MessageDescriptor messageDescriptor, IMessage message, int startRow, int endRow);
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
            var dataRange = ws.Range[ws.Cells[2, 2], ws.Cells[100, fieldCount + 1]];
            dataRange.Clear();
            ws.Cells[2, 2].Value2 = "Loading...";
            ws.Range["A:AZ"].Interior.Color = Color.White;
            dataRange.Interior.Color = Color.LightGray;
        }

        public void PopulateTable<T>(string worksheetName, MessageDescriptor messageDescriptor, RepeatedField<T> responseData)
        {
            Worksheet ws = Application.Sheets[worksheetName];
            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();
            

            var startCol = 2;
            var endCol = fields.Count + 1;
            var startRow = 2;

            var dataRange = ws.Range[ws.Cells[startRow, startCol], ws.Cells[100, endCol]];
            dataRange.Interior.Color = Color.White;

            Range header = ws.Range[ws.Cells[startRow, startCol], ws.Cells[startRow, endCol]];
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
                var rowNumber = rowI + 3;
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
                Range rowRange = ws.Range[ws.Cells[rowNumber, startCol], ws.Cells[rowNumber, endCol]];
                rowRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                rowRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                rowRange.Interior.Color = rowI % 2 == 0 ? Color.LightYellow : Color.White;
            }

            ws.Range["A:AZ"].Columns.AutoFit();
            ws.Range["A:AZ"].Rows.AutoFit();
        }

        public void SetupPaymentRequest()
        {
            Worksheet ws = Application.Sheets[SheetNames.SendPayment];
            ws.Range["A:AZ"].Interior.Color = Color.White;
            Range labelCell = ws.Cells[2, 2];
            labelCell.Value2 = "Payment request:";
            Range inputCell = ws.Cells[2, 3];
            inputCell.Interior.Color = Color.LightYellow;
            inputCell.ColumnWidth = 200;
            ws.Change += WsOnChangeParsePayReq;
            var buttonName = "sendPayment";
            var button = new Microsoft.Office.Tools.Excel.Controls.Button();
            var worksheet = Globals.Factory.GetVstoObject(ws);
            Range selection = ws.Cells[20, 3];
            worksheet.Controls.AddControl(button, selection, buttonName);
            button.Click += ButtonOnClick;
        }

        private void ButtonOnClick(object sender, EventArgs e)
        {
            Worksheet ws = Application.Sheets[SheetNames.SendPayment];
            Range inputCell = ws.Cells[2, 3];
            int responseDataStartRow = 30;
            int responseDataStartColumn = 2;

            string payReq = inputCell.Value2;
            var response = LApp.LndClient.SendPayment(payReq);
            if (response.PaymentError == "")
            {
                PopulateVerticalTable(SheetNames.SendPayment, SendResponse.Descriptor, response, responseDataStartRow);
            }
            else
            {
                Range errorDataLabel = ws.Cells[responseDataStartRow, responseDataStartColumn];
                Range errorData = ws.Cells[responseDataStartRow + 1, responseDataStartColumn];
                errorDataLabel.Value2 = "Payment error:";
                errorData.Value2 = response.PaymentError;
            }
        }

        private void WsOnChangeParsePayReq(Range target)
        {
            if (target.Address == "$C$2")
            {
                PayReq paymentRequest = LApp.LndClient.DecodePaymentRequest(target.Value2);
                PopulateVerticalTable(SheetNames.SendPayment, PayReq.Descriptor, paymentRequest, 5);
            }
        }

        public void MarkAsLoadingVerticalTable(string worksheetName, MessageDescriptor messageDescriptor)
        {
            Worksheet ws = Application.Sheets[worksheetName];
            ws.Select();
            var fieldCount = messageDescriptor.Fields.InDeclarationOrder().Count;
            var dataRange = ws.Range[ws.Cells[2, 2], ws.Cells[fieldCount + 2, 3]];
            dataRange.Clear();
            ws.Cells[2, 2].Value2 = "Loading...";
            ws.Range["A:AZ"].Interior.Color = Color.White;
            dataRange.Interior.Color = Color.LightGray;
            dataRange.Columns.AutoFit();
        }

        public void PopulateVerticalTable(string worksheetName, MessageDescriptor messageDescriptor, IMessage message, int startRow=2, int startColumn=2)
        {
            int endColumn = startColumn + 1;
            Worksheet ws = Application.Sheets[worksheetName];
            ws.Cells[startRow, startColumn].Value2 = "Field Name";
            ws.Cells[startRow, endColumn].Value2 = "Value";
            Range header = ws.Range[ws.Cells[startRow, startColumn], ws.Cells[startRow, endColumn]];
            header.Interior.Color = Color.White;
            header.Font.Bold = true;
            header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();

            int dataStartRow = startRow + 1;
            for (var i = 0; i < fields.Count; i++)
            {
                var field = fields[i];
                var fieldName = field.Name.Replace("_", " ");
                fieldName = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(fieldName);
                int dataRow = dataStartRow + i;
                ws.Cells[dataRow, startColumn].Value2 = fieldName;
                ws.Cells[dataRow, endColumn].Value2 = field.Accessor.GetValue(message).ToString();
                ws.Cells[dataRow, endColumn].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                Range row = ws.Range[ws.Cells[dataRow, startColumn], ws.Cells[dataRow, endColumn]];
                row.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                row.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                row.Interior.Color = i % 2 == 0 ? Color.LightYellow : Color.White;
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
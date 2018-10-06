using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public interface IThisAddIn
    {
        Worksheet SetupSheet(string worksheetName);
        void MarkAsLoadingVerticalTable(string worksheetName, MessageDescriptor messageDescriptor);
        void PopulateVerticalTable(string worksheetName, MessageDescriptor messageDescriptor, IMessage message);
    }

    public partial class ThisAddIn: IThisAddIn
    {

        public LightningApp LApp;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.WorkbookOpen += ApplicationOnWorkbookOpen;
        }
        
        // Check to see if the workbook is an LNDExcel workbook
        private void ApplicationOnWorkbookOpen(Workbook wb)
        {
            var isLndWorkbook = false;
            try
            {
                Worksheet ws = Application.Sheets[SheetNames.GetInfo];
                if (ws.Cells[1, 1].Value2 == "LNDExcel")
                {
                    isLndWorkbook = true;
                }
            }
            catch (COMException)
            {
                // GetInfo tab doesn't exist, certainly not an LNDExcel workbook
            }

            if (isLndWorkbook)
            {
                ConnectLnd();
            }
        }

        public void ConnectLnd()
        {
            this.LApp = new LightningApp(this);
            MarkLndExcelWorkbook();
            LApp.RefreshGetInfo();
            Application.ActiveWorkbook.SheetActivate += Workbook_SheetActivate;
        }

        public Worksheet SetupSheet(string worksheetName)
        {
            Worksheet ws;
            try
            {
                ws = Application.Sheets[worksheetName];
            }
            catch (COMException)
            {
                Application.Sheets.Add();
                ws = Application.ActiveSheet;
                ws.Name = worksheetName;
                LApp.RefreshGetInfo();
            }

            return ws;
        }
        
        private void Workbook_SheetActivate(object sh)
        {
            var ws = (Worksheet) sh;
            switch (ws.Name)
            {
                case SheetNames.GetInfo:
                    LApp.RefreshGetInfo();
                    break;
                case SheetNames.Channels:
                    LApp.RefreshChannels();
                    break;
                case SheetNames.Payments:
                    LApp.RefreshPayments();
                    break;
            }
        }

        public void MarkLndExcelWorkbook()
        {
            Worksheet ws = SetupSheet(SheetNames.GetInfo);
            ws.Cells[1, 1].Value2 = "LNDExcel";
            ws.Cells[1, 1].Font.Color = Color.White;
        }

        public void MarkAsLoadingTable(string worksheetName, MessageDescriptor messageDescriptor)
        {
            var ws = SetupSheet(worksheetName);
            var fieldCount = messageDescriptor.Fields.InDeclarationOrder().Count;
            var dataRange = ws.Range[ws.Cells[2, 2], ws.Cells[3, fieldCount + 1]];
            dataRange.Clear();
            ws.Cells[2, 2].Value2 = "Loading...";
            ws.Range["A:AZ"].Interior.Color = Color.White;
            dataRange.Interior.Color = Color.LightGray;
        }

        public void PopulateTable<T>(string worksheetName, MessageDescriptor messageDescriptor, RepeatedField<T> responseData)
        {
            Worksheet ws = SetupSheet(worksheetName);
            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();

            var startCol = 2;
            var endCol = fields.Count + 1;
            var startRow = 2;

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
            }

            for (var rowI = 0; rowI < responseData.Count; rowI++)
            {
                T data = responseData[rowI];
                var rowNumber = rowI + 3;
                for (var colJ = 0; colJ < fields.Count; colJ++)
                {
                    var colNumber = colJ + 2;
                    var dataCell = ws.Cells[rowNumber, colNumber];
                    dataCell.Value2 = fields[colJ].Accessor.GetValue(data as IMessage).ToString();
                    dataCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    dataCell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                    dataCell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    dataCell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
                    dataCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                }
                Range rowRange = ws.Range[ws.Cells[rowNumber, startCol], ws.Cells[rowNumber, endCol]];
                rowRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                rowRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                rowRange.Interior.Color = rowI % 2 == 0 ? Color.LightYellow : Color.White;
            }

            ws.Range["A:AZ"].Columns.AutoFit();
        }

        public void MarkAsLoadingVerticalTable(string worksheetName, MessageDescriptor messageDescriptor)
        {
            Worksheet ws = SetupSheet(worksheetName);
            ws.Select();
            var fieldCount = messageDescriptor.Fields.InDeclarationOrder().Count;
            var dataRange = ws.Range[ws.Cells[2, 2], ws.Cells[fieldCount + 2, 3]];
            dataRange.Clear();
            ws.Cells[2, 2].Value2 = "Loading...";
            ws.Range["A:AZ"].Interior.Color = Color.White;
            dataRange.Interior.Color = Color.LightGray;
            dataRange.Columns.AutoFit();
        }

        public void PopulateVerticalTable(string worksheetName, MessageDescriptor messageDescriptor, IMessage message)
        {
            Worksheet ws = SetupSheet(worksheetName);
            ws.Cells[2, 2].Value2 = "Field Name";
            ws.Cells[2, 3].Value2 = "Value";
            Range header = ws.Range[ws.Cells[2, 2], ws.Cells[2, 3]];
            header.Interior.Color = Color.White;
            header.Font.Bold = true;
            header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();

            for (var i = 0; i < fields.Count; i++)
            {
                var field = fields[i];
                var fieldName = field.Name.Replace("_", " ");
                fieldName = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(fieldName);

                ws.Cells[i + 3, 2].Value2 = fieldName;
                ws.Cells[i + 3, 3].Value2 = field.Accessor.GetValue(message).ToString();
                ws.Cells[i + 3, 3].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                Range row = ws.Range[ws.Cells[i + 3, 2], ws.Cells[i + 3, 3]];
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
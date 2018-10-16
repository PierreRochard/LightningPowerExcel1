using System.Collections.Generic;
using System.Drawing;
using System.Threading;

using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;

using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class Tables
    {
        public static void MarkAsLoadingTable(Worksheet ws, MessageDescriptor messageDescriptor)
        {
            var fieldCount = messageDescriptor.Fields.InDeclarationOrder().Count;
            var dataRange = ws.Range[ws.Cells[3, 2], ws.Cells[100, fieldCount]];
            dataRange.Clear();
            ws.Cells[3, 2].Value2 = "Loading...";
            dataRange.Interior.Color = Color.LightGray;
        }

        public static void ClearTable(Worksheet ws, MessageDescriptor messageDescriptor, int startRow = 2,
            int startColumn = 2)
        {
            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();
            var endColumn = fields.Count + 1;

            // Skip the title
            startRow++;

            // Skip the table headers
            startRow++;

            // Find the last row in the table
            var endRow = startRow;
            Range lastCell = ws.Cells[endRow, startColumn];
            while (!string.IsNullOrWhiteSpace(lastCell.Value2))
            {
                endRow++;
                lastCell = ws.Cells[endRow, startColumn];
            }

            var dataRange = ws.Range[ws.Cells[startRow, startColumn], ws.Cells[endRow, endColumn]];
            dataRange.Clear();
        }

        public static void PopulateTable<T>(Worksheet ws, MessageDescriptor messageDescriptor, RepeatedField<T> responseData, int startRow = 2, int startColumn = 2)
        {
            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();

            // Skip title
            startRow++;
            
            // Skip header
            startRow++;

            for (var rowI = 0; rowI < responseData.Count; rowI++)
            {
                var rowNumber = rowI + startRow;
                for (var colJ = 0; colJ < fields.Count; colJ++)
                {
                    var field = fields[colJ];
                    var colNumber = colJ + 2;
                    var dataCell = ws.Cells[rowNumber, colNumber];

                    string value = "";
                    
                    T data = responseData[rowI];
                    if (field.IsRepeated && field.FieldType != FieldType.Message)
                    {
                        var items = (RepeatedField<string>)fields[colJ].Accessor.GetValue(data as IMessage);
                        for (int i = 0; i < items.Count; i++)
                        {
                            value += items[i];
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
                }
            }
        }

        public static void SetupTable<T>(Worksheet ws, string tableTitle, MessageDescriptor messageDescriptor, RepeatedField<T> responseData = null, int startRow = 2, int startColumn = 2)
        {
            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();

            var endCol = fields.Count + 1;

            Range title = ws.Cells[startRow, startColumn];
            title.Font.Italic = true;
            title.Value2 = tableTitle;

            startRow++;

            var dataRange = ws.Range[ws.Cells[startRow, startColumn], ws.Cells[100, endCol]];
            dataRange.Clear();
            Formatting.RemoveBorders(dataRange);
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

            int rowCount = responseData?.Count ?? 2;
            for (var rowI = 0; rowI < rowCount; rowI++)
            {

                var rowNumber = rowI + startRow + 1;
                for (var colJ = 0; colJ < fields.Count; colJ++)
                {
                    var field = fields[colJ];
                    var colNumber = colJ + 2;
                    var dataCell = ws.Cells[rowNumber, colNumber];

                    string value = "";
                    if (responseData != null)
                    {
                        T data = responseData[rowI];
                        if (field.IsRepeated && field.FieldType != FieldType.Message)
                        {
                            var items = (RepeatedField<string>)fields[colJ].Accessor.GetValue(data as IMessage);
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

        public static void MarkAsLoadingVerticalTable(Worksheet ws, MessageDescriptor messageDescriptor)
        {
            ws.Select();
            var fieldCount = messageDescriptor.Fields.InDeclarationOrder().Count;
            var dataRange = ws.Range[ws.Cells[3, 2], ws.Cells[fieldCount, 3]];
            dataRange.Clear();
            ws.Cells[3, 2].Value2 = "Loading...";
            dataRange.Interior.Color = Color.LightGray;
            dataRange.Columns.AutoFit();
        }

        public static void SetupVerticalTable(Worksheet ws, string tableTitle, MessageDescriptor messageDescriptor, IMessage message = null, int startRow = 2, int startColumn = 2)
        {
            int endColumn = startColumn + 1;

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

        public static void PopulateVerticalTable(Worksheet ws, MessageDescriptor messageDescriptor, IMessage message, int startRow = 2, int startColumn = 2)
        {
            int endColumn = startColumn + 1;

            // Skip title
            startRow++;

            int dataRow = startRow;
            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();
            foreach (var field in fields)
            {
                if (field.IsRepeated)
                {
                    continue;
                }
                ws.Cells[dataRow, endColumn].Value2 = message != null ? field.Accessor.GetValue(message).ToString() : "";
                
                dataRow++;
            }
        }

        public static void ClearVerticalTable(Worksheet ws, MessageDescriptor messageDescriptor, int startRow = 2,
            int startColumn = 2)
        {
            // Skip the title
            startRow++;

            IList<FieldDescriptor> fields = messageDescriptor.Fields.InDeclarationOrder();
            var endRow = startRow + fields.Count;
            var dataColumn = startColumn + 1;
            var dataRange = ws.Range[ws.Cells[startRow, dataColumn], ws.Cells[endRow, dataColumn]];
            dataRange.Clear();
        }
    }
}
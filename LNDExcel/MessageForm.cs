using System;
using System.Collections.Generic;
using Google.Protobuf;
using Google.Protobuf.Reflection;
using Microsoft.Office.Interop.Excel;
using Button = Microsoft.Office.Tools.Excel.Controls.Button;

namespace LNDExcel
{
    public class MessageForm<TRequestMessage> where TRequestMessage: IMessage
    {
        public int StartRow;
        public int StartColumn;
        public int EndRow;
        public int EndColumn;
        public IList<FieldDescriptor> Fields;
        private int _dataStartRow;
        private Range _errorData;

        public MessageForm(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor descriptor, string title, int startRow = 2,
            int startColumn = 2)
        {
            Fields = descriptor.Fields.InDeclarationOrder();
            StartRow = startRow;
            EndRow = StartRow + Fields.Count;
            StartColumn = startColumn;
            EndColumn = StartColumn + 1;


            var titleCell = ws.Cells[StartRow, StartColumn];
            titleCell.Font.Italic = true;
            titleCell.Value2 = title;
            _dataStartRow = StartRow + 1;

            var form = ws.Range[ws.Cells[_dataStartRow, StartColumn], ws.Cells[EndRow, EndColumn]];
            Formatting.VerticalTable(form);

            var header = ws.Range[ws.Cells[_dataStartRow, StartColumn], ws.Cells[EndRow, StartColumn]];
            Formatting.VerticalTableHeaderColumn(header);

            var data = ws.Range[ws.Cells[_dataStartRow, EndColumn], ws.Cells[EndRow, EndColumn]];
            Formatting.VerticalTableDataColumn(data);

            var rowIndex = 0;
            foreach (var field in Fields)
            {

                var rowNumber = _dataStartRow + rowIndex;

                var headerCell = ws.Cells[rowNumber, StartColumn];
                
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;

                var rowRange = ws.Range[ws.Cells[rowNumber, StartColumn], ws.Cells[rowNumber, EndColumn]];
                Formatting.VerticalTableRow(rowRange, rowNumber);

                rowIndex++;
            }
            
            var submitButtonRow = rowIndex + 2;
            Button submitButton = Utilities.CreateButton("submit", ws, ws.Cells[submitButtonRow, StartColumn], "Submit");
            submitButton.Click += SubmitButtonOnClick;
            _errorData = ws.Cells[submitButtonRow, StartColumn + 1];



            titleCell.Columns.AutoFit();
        }

        private void SubmitButtonOnClick(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
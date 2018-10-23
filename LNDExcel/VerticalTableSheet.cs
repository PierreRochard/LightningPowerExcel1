using System.Collections.Generic;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class VerticalTableSheet<TMessageClass> where TMessageClass : IMessage
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        private int _startRow;
        private int _dataStartRow;
        private int _startColumn;
        private int _endColumn;
        private int _endRow;

        private TMessageClass _data;
        private readonly IList<FieldDescriptor> _fields;

        public VerticalTableSheet(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor messageDescriptor)
        {
            Ws = ws;
            LApp = lApp;
            _fields = messageDescriptor.Fields.InDeclarationOrder();
        }
        
        public void SetupVerticalTable(string tableName, int startRow = 2, int startColumn = 2)
        {
            _startRow = startRow;
            _dataStartRow = startRow + 1;
            _startColumn = startColumn;
            _endColumn = _startColumn + 1;

            _endRow = startRow + _fields.Count;

            var title = Ws.Cells[_startRow, _startColumn];
            title.Font.Italic = true;
            title.Value2 = tableName;

            var table = Ws.Range[Ws.Cells[_dataStartRow, _startColumn], Ws.Cells[_endRow, _endColumn]];
            Formatting.VerticalTable(table);

            var header = Ws.Range[Ws.Cells[_dataStartRow, _startColumn], Ws.Cells[_endRow, _startColumn]];
            Formatting.VerticalTableHeaderColumn(header);

            var data = Ws.Range[Ws.Cells[_dataStartRow, _endColumn], Ws.Cells[_endRow, _endColumn]];
            Formatting.VerticalTableDataColumn(data);

            for (var fieldIndex = 0; fieldIndex < _fields.Count; fieldIndex++)
            {
                var rowNumber = _dataStartRow + fieldIndex;
                var headerCell = Ws.Cells[rowNumber, _startColumn];
                var field = _fields[fieldIndex];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;
            }

            for (var rowI = 0; rowI < _fields.Count; rowI++)
            {
                var rowNumber = _dataStartRow + rowI;
                var rowRange = Ws.Range[Ws.Cells[rowNumber, _startColumn], Ws.Cells[rowNumber, _endColumn]];
                Formatting.VerticalTableRow(rowRange, rowNumber % 2 == 0);
            }
        }
        
        public void Update(TMessageClass newMessage)
        {
            
            var isCached = _data != null;
            if (isCached && _data.Equals(newMessage))
            {
                return;
            }


            if (!isCached)
            {
                Populate(newMessage);
            }
            else
            {
                Update(newMessage, _data);
            }

            Ws.Range["A:C"].Columns.AutoFit();
            Ws.Range["A:C"].Rows.AutoFit();
            _data = newMessage;
        }

        private void Populate(TMessageClass newMessage)
        {
            for (var fieldIndex = 0; fieldIndex < _fields.Count; fieldIndex++)
            {
                var field = _fields[fieldIndex];
                var dataCell = Ws.Cells[_dataStartRow + fieldIndex, _endColumn];
                var value = "";

                if (field.IsRepeated && field.FieldType != FieldType.Message)
                {
                    var items = (RepeatedField<string>)field.Accessor.GetValue(newMessage);
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
                    value = field.Accessor.GetValue(newMessage).ToString();
                }
                dataCell.Value2 = value;
            }
        }

        public void Update(TMessageClass newMessage, TMessageClass oldMessage)
        {
            for (var fieldIndex = 0; fieldIndex < _fields.Count; fieldIndex++)
            {
                var field = _fields[fieldIndex];
                var newValue = field.Accessor.GetValue(newMessage).ToString();
                var oldValue = field.Accessor.GetValue(oldMessage).ToString();
                if (oldValue == newValue) continue;

                var dataCell = Ws.Cells[_dataStartRow + fieldIndex, _endColumn];
                var value = "";

                if (field.IsRepeated && field.FieldType != FieldType.Message)
                {
                    var items = (RepeatedField<string>)field.Accessor.GetValue(newMessage);
                    for (var i = 0; i < items.Count; i++)
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
                    value = newValue;
                }

                dataCell.Value2 = value;
            }

        }
    }
}
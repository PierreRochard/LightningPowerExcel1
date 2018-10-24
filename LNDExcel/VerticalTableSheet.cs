using System.Collections.Generic;
using System.Linq;
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
        private readonly IReadOnlyCollection<string> _excludeList;

        public VerticalTableSheet(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor messageDescriptor, 
            IReadOnlyCollection<string> excludeList = default(List<string>))
        {
            Ws = ws;
            LApp = lApp;
            _fields = messageDescriptor.Fields.InDeclarationOrder();
            _excludeList = excludeList;
        }
        
        public void SetupVerticalTable(string tableName, int startRow = 2, int startColumn = 2)
        {
            _startRow = startRow;
            _dataStartRow = startRow + 1;
            _startColumn = startColumn;
            _endColumn = _startColumn + 1;

            if (_excludeList == null)
            {
                _endRow = startRow + _fields.Count;
            }
            else
            {
                _endRow = startRow + _fields.Count(f => !_excludeList.Any(f.Name.Contains));
            }

            var title = Ws.Cells[_startRow, _startColumn];
            title.Font.Italic = true;
            title.Value2 = tableName;

            var table = Ws.Range[Ws.Cells[_dataStartRow, _startColumn], Ws.Cells[_endRow, _endColumn]];
            Formatting.VerticalTable(table);

            var header = Ws.Range[Ws.Cells[_dataStartRow, _startColumn], Ws.Cells[_endRow, _startColumn]];
            Formatting.VerticalTableHeaderColumn(header);

            var data = Ws.Range[Ws.Cells[_dataStartRow, _endColumn], Ws.Cells[_endRow, _endColumn]];
            Formatting.VerticalTableDataColumn(data);

            var rowIndex = 0;
            foreach (var field in _fields)
            {
                if (_excludeList != null && _excludeList.Any(field.Name.Contains)) continue;

                var rowNumber = _dataStartRow + rowIndex;

                var headerCell = Ws.Cells[rowNumber, _startColumn];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;

                var rowRange = Ws.Range[Ws.Cells[rowNumber, _startColumn], Ws.Cells[rowNumber, _endColumn]];
                Formatting.VerticalTableRow(rowRange, rowNumber % 2 == 0);

                rowIndex++;
            }
        }

        public void Clear()
        {
            var data = Ws.Range[Ws.Cells[_dataStartRow, _endColumn], Ws.Cells[_endRow, _endColumn]];
            data.ClearContents();
            _data = default(TMessageClass);
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
        }

        public void Populate(TMessageClass newMessage)
        {
            var rowIndex = 0;
            foreach (var field in _fields)
            {
                if (_excludeList != null && _excludeList.Any(field.Name.Contains)) continue;

                var dataCell = Ws.Cells[_dataStartRow + rowIndex, _endColumn];
                var value = string.Empty;
                if (field.IsRepeated && field.Accessor.GetValue(newMessage) is RepeatedField<object> items)
                {
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
                    value = field.Accessor.GetValue(newMessage).ToString();
                }
                dataCell.Value2 = value;
                rowIndex++;
            }
            _data = newMessage;
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
            _data = newMessage;

        }
    }
}
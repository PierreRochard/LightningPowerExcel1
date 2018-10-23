using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class TableSheet<TMessageClass> where TMessageClass : IMessage
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        private int _startRow;
        private int _headerRow;
        private int _startColumn;
        private int _endColumn;

        private readonly Dictionary<object, TMessageClass> _data;
        private readonly MessageDescriptor _messageDescriptor;
        private readonly IFieldAccessor _uniqueKeyField;
        private readonly string _uniqueKeyName;

        public TableSheet(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor messageDescriptor, string uniqueKeyName)
        {
            Ws = ws;
            LApp = lApp;
            _data = new Dictionary<object, TMessageClass>();
            _messageDescriptor = messageDescriptor;
            _uniqueKeyName = uniqueKeyName;
            var fields = messageDescriptor.Fields.InDeclarationOrder();
            _uniqueKeyField = fields.First(m => m.Name == _uniqueKeyName).Accessor;
        }

        public void SetupTable(string tableName, int rowCount, int startRow = 2, int startColumn = 2)
        {
            _startRow = startRow;
            _headerRow = _startRow + 1;
            _startColumn = startColumn;
            _endColumn = _startColumn + _messageDescriptor.Fields.InFieldNumberOrder().Count - 1;

            var fields = _messageDescriptor.Fields.InDeclarationOrder();

            var title = Ws.Cells[_startRow, _startColumn];
            title.Font.Italic = true;
            title.Value2 = tableName;

            var header = Ws.Range[Ws.Cells[_headerRow, _startColumn], Ws.Cells[_headerRow, _endColumn]];
            Formatting.TableHeaderRow(header);

            for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var columnNumber = _startColumn + fieldIndex;
                var headerCell = Ws.Cells[_headerRow, columnNumber];
                var field = fields[fieldIndex];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;
                if (field.IsRepeated && field.FieldType != FieldType.Message)
                {
                    Ws.Columns[columnNumber].ColumnWidth = 100;
                }
            }

            for (var rowI = 1; rowI <= rowCount; rowI++)
            {
                var rowNumber = rowI + _headerRow;
                var rowRange = Ws.Range[Ws.Cells[rowNumber, _startColumn], Ws.Cells[rowNumber, _endColumn]];
                Formatting.TableDataRow(rowRange, rowNumber % 2 == 0);
            }
        }

        public void Update(RepeatedField<TMessageClass> data)
        {
            foreach (var newMessage in data)
            {
                var uniqueKey = _uniqueKeyField.GetValue(newMessage);
                var isCached = _data.TryGetValue(uniqueKey, out var cachedMessage);
                if (isCached && cachedMessage.Equals(newMessage))
                {
                    continue;
                }

                _data[uniqueKey] = newMessage;

                if (!isCached)
                {
                    AppendRow(newMessage);
                }
                else
                {
                    UpdateRow(newMessage, cachedMessage);
                }
            }

            foreach (var cachedUniqueKey in _data.Keys)
            {
                var result = data.FirstOrDefault(newMessage => _uniqueKeyField.GetValue(newMessage).ToString() == cachedUniqueKey.ToString());
                if (result == null)
                {
                    RemoveRow(cachedUniqueKey);
                }
            }

            Ws.Range["A:AZ"].Columns.AutoFit();
            Ws.Range["A:AZ"].Rows.AutoFit();
        }

        private void RemoveRow(object uniqueKey)
        {
            var rowNumber = GetRow(uniqueKey);
            if (rowNumber == 0) return;
            var range = Ws.Range[Ws.Cells[rowNumber, _startColumn], Ws.Cells[rowNumber, _endColumn]];
            range.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        private void AppendRow(TMessageClass newMessage)
        {
            var lastRow = GetLastRow();
            Formatting.TableDataRow(Ws.Range[Ws.Cells[lastRow, _startColumn], Ws.Cells[lastRow, _endColumn]], lastRow % 2 == 0);
            var fields = _messageDescriptor.Fields.InDeclarationOrder();
            for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var field = fields[fieldIndex];
                var dataCell = Ws.Cells[lastRow, _startColumn + fieldIndex];
                var newValue = field.Accessor.GetValue(newMessage).ToString();
                AssignCellValue(newMessage, field, newValue, dataCell);
            }
        }

        public void UpdateRow(TMessageClass newMessage, TMessageClass oldMessage)
        {
            var row = GetRow(_uniqueKeyField.GetValue(newMessage));
            if (row == 0)
            {
                AppendRow(newMessage);
                return;
            }

            var fields = _messageDescriptor.Fields.InDeclarationOrder();
            for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var field = fields[fieldIndex];
                var newValue = field.Accessor.GetValue(newMessage).ToString();
                var oldValue = field.Accessor.GetValue(oldMessage).ToString();
                if (oldValue == newValue) continue;

                var dataCell = Ws.Cells[row, _startColumn + fieldIndex];
                AssignCellValue(newMessage, field, newValue, dataCell);

            }

        }

        private static void AssignCellValue(TMessageClass newMessage, FieldDescriptor field, string newValue, dynamic dataCell)
        {
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
                dataCell.Value2 = value;
            }
            else if (field.FieldType == FieldType.UInt64)
            {
                dataCell.NumberFormat = "@";
                dataCell.Value2 = newValue;
            }
            else
            {
                dataCell.Value2 = newValue;
            }
        }

        private int GetRow(object uniqueKey)
        {
            var uniqueKeyString = uniqueKey.ToString();
            var idColumn = 1;
            Range idColumnNameCell = Ws.Cells[_headerRow, idColumn];
            var uniqueKeyName = Utilities.FormatFieldName(_uniqueKeyName);
            while (idColumnNameCell.Value2 == null || idColumnNameCell.Value2.ToString() != uniqueKeyName)
            {
                idColumn++;
                idColumnNameCell = Ws.Cells[_headerRow, idColumn];
            }

            var rowNumber = _headerRow;
            var uniqueKeyCellString = UniqueKeyCellString(Ws.Cells[rowNumber, idColumn]);
            while (uniqueKeyCellString != uniqueKeyString)
            {
                rowNumber++;
                if (rowNumber > 100)
                {
                    return 0;
                }
                uniqueKeyCellString = UniqueKeyCellString(Ws.Cells[rowNumber, idColumn]);
            }

            return rowNumber;
        }

        private static string UniqueKeyCellString(Range uniqueKeyCell)
        {
            if (uniqueKeyCell.Value2 == null) return string.Empty;
            var uniqueKeyCellString = uniqueKeyCell.Value2.ToString();
            return uniqueKeyCellString;
        }

        private int GetLastRow()
        {
            var lastRow = _headerRow + 1;
            Range dataCell = Ws.Cells[lastRow, _startColumn];
            while (dataCell.Value2 != null && !string.IsNullOrWhiteSpace(dataCell.Value2.ToString()))
            {
                lastRow++;
                dataCell = Ws.Cells[lastRow, _startColumn];
            }

            return lastRow;

        }
    }
}

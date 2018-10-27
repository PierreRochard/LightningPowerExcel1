using System.Collections.Generic;
using System.Linq;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class TableSheet<TMessageClass> where TMessageClass : IMessage
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        private int _startRow;
        private int _headerRow;
        public int StartColumn;
        public int EndColumn;
        public IList<FieldDescriptor> _fields;

        public Dictionary<object, TMessageClass> Data;
        public RepeatedField<TMessageClass> DataList;
        private readonly MessageDescriptor _messageDescriptor;
        private readonly IFieldAccessor _uniqueKeyField;
        private readonly string _uniqueKeyName;

        private readonly List<string> _wideColumns;

        public TableSheet(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor messageDescriptor, string uniqueKeyName,
            List<string> wideColumns = null)
        {
            Ws = ws;
            LApp = lApp;
            Data = new Dictionary<object, TMessageClass>();
            _messageDescriptor = messageDescriptor;
            _uniqueKeyName = uniqueKeyName;
            _fields = messageDescriptor.Fields.InDeclarationOrder();
            _uniqueKeyField = _fields.First(m => m.Name == _uniqueKeyName).Accessor;
            _wideColumns = wideColumns;
        }

        public void SetupTable(string tableName, int rowCount, int startRow = 2, int startColumn = 2)
        {
            _startRow = startRow;
            _headerRow = _startRow + 1;
            StartColumn = startColumn;
            EndColumn = StartColumn + _messageDescriptor.Fields.InFieldNumberOrder().Count - 1;


            var title = Ws.Cells[_startRow, StartColumn];
            title.Font.Italic = true;
            title.Value2 = tableName;

            var header = Ws.Range[Ws.Cells[_headerRow, StartColumn], Ws.Cells[_headerRow, EndColumn]];
            Formatting.TableHeaderRow(header);
            
            Ws.Columns.AutoFit();
            for (var fieldIndex = 0; fieldIndex < _fields.Count; fieldIndex++)
            {
                var columnNumber = StartColumn + fieldIndex;
                var headerCell = Ws.Cells[_headerRow, columnNumber];
                var field = _fields[fieldIndex];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;

                if (field.IsRepeated && field.FieldType != FieldType.Message)
                {
                    Ws.Columns[columnNumber].ColumnWidth = 100;
                }

                var isWide = _wideColumns != null && _wideColumns.Any(field.Name.Contains);
                if (!isWide) continue;
                Formatting.WideTableColumn(Ws.Range[Ws.Cells[1, StartColumn], Ws.Cells[100, EndColumn]]);
            }

            for (var rowI = 1; rowI <= rowCount; rowI++)
            {
                var rowNumber = rowI + _headerRow;
                var rowRange = Ws.Range[Ws.Cells[rowNumber, StartColumn], Ws.Cells[rowNumber, EndColumn]];
                Formatting.TableDataRow(rowRange, rowNumber % 2 == 0);
            }


        }

        public void Update(RepeatedField<TMessageClass> data)
        {
            DataList = data;
            foreach (var newMessage in DataList)
            {
                var uniqueKey = _uniqueKeyField.GetValue(newMessage);
                var isCached = Data.TryGetValue(uniqueKey, out var cachedMessage);
                if (isCached && cachedMessage.Equals(newMessage))
                {
                    continue;
                }

                Data[uniqueKey] = newMessage;

                if (!isCached)
                {
                    AppendRow(newMessage);
                }
                else
                {
                    UpdateRow(newMessage, cachedMessage);
                }
            }

            foreach (var cachedUniqueKey in Data.Keys)
            {
                var result = DataList.FirstOrDefault(newMessage => _uniqueKeyField.GetValue(newMessage).ToString() == cachedUniqueKey.ToString());
                if (result == null)
                {
                    RemoveRow(cachedUniqueKey);
                }
            }

            Ws.Columns.AutoFit();
            for (var fieldIndex = 0; fieldIndex < _fields.Count; fieldIndex++)
            {
                var field = _fields[fieldIndex];
                var columnNumber = StartColumn + fieldIndex;
                var isWide = _wideColumns != null && _wideColumns.Any(field.Name.Contains);
                if (!isWide) continue;
                Formatting.WideTableColumn(Ws.Range[Ws.Cells[1, columnNumber], Ws.Cells[1, columnNumber]]);
            }
        }

        private void RemoveRow(object uniqueKey)
        {
            var rowNumber = GetRow(uniqueKey);
            if (rowNumber == 0) return;
            var range = Ws.Range[Ws.Cells[rowNumber, StartColumn], Ws.Cells[rowNumber, EndColumn]];
            range.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        private void AppendRow(TMessageClass newMessage)
        {
            var lastRow = GetLastRow();
            Formatting.TableDataRow(Ws.Range[Ws.Cells[lastRow, StartColumn], Ws.Cells[lastRow, EndColumn]], lastRow % 2 == 0);
            var fields = _messageDescriptor.Fields.InDeclarationOrder();
            for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var field = fields[fieldIndex];
                var columnNumber = StartColumn + fieldIndex;
                var dataCell = Ws.Cells[lastRow, columnNumber];
                var newValue = field.Accessor.GetValue(newMessage).ToString();
                AssignCellValue(newMessage, field, newValue, dataCell);
                var isWide = _wideColumns != null && _wideColumns.Any(field.Name.Contains);
                Formatting.TableDataColumn(Ws.Range[Ws.Cells[lastRow, columnNumber], Ws.Cells[lastRow, columnNumber]], isWide);
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

                var dataCell = Ws.Cells[row, StartColumn + fieldIndex];
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
            Range dataCell = Ws.Cells[lastRow, StartColumn];
            while (dataCell.Value2 != null && !string.IsNullOrWhiteSpace(dataCell.Value2.ToString()))
            {
                lastRow++;
                dataCell = Ws.Cells[lastRow, StartColumn];
            }
            return lastRow;
        }

        public void Clear()
        {
            var data = Ws.Range[Ws.Cells[_headerRow + 1, StartColumn], Ws.Cells[GetLastRow(), EndColumn]];
            data.ClearContents();
            Data = new Dictionary<object, TMessageClass>();
        }
    }
}

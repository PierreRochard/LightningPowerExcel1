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
        
        public int StartRow;
        public int HeaderRow;
        public int StartColumn;
        public int EndColumn;
        public int EndRow;

        public IList<FieldDescriptor> Fields;
        public Dictionary<object, TMessageClass> Data;
        public RepeatedField<TMessageClass> DataList;
        public Range Title;

        private readonly List<string> _wideColumns;
        private readonly IFieldAccessor _uniqueKeyField;
        private readonly string _uniqueKeyName;

        public TableSheet(Worksheet ws, AsyncLightningApp lApp, MessageDescriptor messageDescriptor, string uniqueKeyName,
            List<string> wideColumns = null, bool nestedData = false)
        {
            Ws = ws;
            LApp = lApp;
            Data = new Dictionary<object, TMessageClass>();
            Fields = messageDescriptor.Fields.InDeclarationOrder()
                .Where(f => f.FieldType != FieldType.Message || !nestedData).ToList();
            
            var nestedFields = messageDescriptor.Fields.InDeclarationOrder()
                .Where(f => f.FieldType == FieldType.Message && nestedData).ToList();
            foreach (var field in nestedFields)
            {
                Fields = Fields.Concat(field.MessageType.Fields.InDeclarationOrder()).ToArray();
            }

            _uniqueKeyName = uniqueKeyName;
            _uniqueKeyField = Fields.First(m => m.Name == _uniqueKeyName).Accessor;
            _wideColumns = wideColumns;
        }

        public void SetupTable(string tableName, int rowCount, int startRow = 2, int startColumn = 2)
        {
            StartRow = startRow;
            HeaderRow = startRow + 1;
            StartColumn = startColumn;
            EndColumn = StartColumn + Fields.Count - 1;
            EndRow = HeaderRow + rowCount;


            Title = Ws.Cells[StartRow, StartColumn];
            Title.Value2 = tableName;

            var header = Ws.Range[Ws.Cells[HeaderRow, StartColumn], Ws.Cells[HeaderRow, EndColumn]];
            Formatting.TableHeaderRow(header);

            var data = Ws.Range[Ws.Cells[HeaderRow + 1, StartColumn], Ws.Cells[EndRow, EndColumn]];
            Formatting.TableDataColumn(data, false);

            Ws.Columns.AutoFit();
            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var columnNumber = StartColumn + fieldIndex;
                var headerCell = Ws.Cells[HeaderRow, columnNumber];
                var field = Fields[fieldIndex];
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
                var rowNumber = rowI + HeaderRow;
                var rowRange = Ws.Range[Ws.Cells[rowNumber, StartColumn], Ws.Cells[rowNumber, EndColumn]];
                Formatting.TableDataRow(rowRange, rowNumber % 2 == 0);
            }

            Formatting.TableTitle(Title);
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
            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var field = Fields[fieldIndex];
                var columnNumber = StartColumn + fieldIndex;
                var isWide = _wideColumns != null && _wideColumns.Any(field.Name.Contains);
                if (!isWide) continue;
                Formatting.WideTableColumn(Ws.Range[Ws.Cells[1, columnNumber], Ws.Cells[1, columnNumber]]);
            }

            Formatting.TableTitle(Title);
            EndRow = GetLastRow();
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
            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var field = Fields[fieldIndex];
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

            for (var fieldIndex = 0; fieldIndex < Fields.Count; fieldIndex++)
            {
                var field = Fields[fieldIndex];
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
            Range idColumnNameCell = Ws.Cells[HeaderRow, idColumn];
            var uniqueKeyName = Utilities.FormatFieldName(_uniqueKeyName);
            while (idColumnNameCell.Value2 == null || idColumnNameCell.Value2.ToString() != uniqueKeyName)
            {
                idColumn++;
                idColumnNameCell = Ws.Cells[HeaderRow, idColumn];
            }

            var rowNumber = HeaderRow;
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
            var lastRow = HeaderRow + 1;
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
            var data = Ws.Range[Ws.Cells[HeaderRow + 1, StartColumn], Ws.Cells[GetLastRow(), EndColumn]];
            data.ClearContents();
            Data = new Dictionary<object, TMessageClass>();
        }
    }
}

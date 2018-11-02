using System;
using System.Collections.Generic;
using Google.Protobuf;
using Google.Protobuf.Reflection;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Button = Microsoft.Office.Tools.Excel.Controls.Button;

namespace LNDExcel
{
    public class MessageForm<TRequestMessage, TResponseMessage> where TRequestMessage: IMessage, new()
        where TResponseMessage: IMessage
    {
        public Worksheet Ws;
        private readonly AsyncLightningApp _lApp;
        private readonly Func<TRequestMessage, TResponseMessage> _query;
        public int StartRow;
        public int StartColumn;
        public int EndRow;
        public int EndColumn;
        public IList<FieldDescriptor> Fields;
        public Range ErrorData;
        private readonly int _dataStartRow;
        
        public MessageForm(Worksheet ws, AsyncLightningApp lApp, Func<TRequestMessage, TResponseMessage> query, MessageDescriptor descriptor, string title, int startRow = 2,
            int startColumn = 2)
        {
            Ws = ws;
            _lApp = lApp;
            _query = query;
            Fields = descriptor.Fields.InDeclarationOrder();
            StartRow = startRow;
            StartColumn = startColumn;
            EndColumn = StartColumn + 1;

            var titleCell = ws.Cells[StartRow, StartColumn];
            titleCell.Font.Italic = true;
            titleCell.Value2 = title;

            _dataStartRow = StartRow + 1;
            var endDataRow = startRow + Fields.Count;

            var form = ws.Range[ws.Cells[_dataStartRow, StartColumn], ws.Cells[endDataRow, EndColumn]];
            Formatting.VerticalTable(form);

            var header = ws.Range[ws.Cells[_dataStartRow, StartColumn], ws.Cells[endDataRow, StartColumn]];
            Formatting.VerticalTableHeaderColumn(header);

            var data = ws.Range[ws.Cells[_dataStartRow, EndColumn], ws.Cells[endDataRow, EndColumn]];
            Formatting.VerticalTableDataColumn(data);

            var rowNumber = _dataStartRow;
            foreach (var field in Fields)
            {
                var headerCell = ws.Cells[rowNumber, StartColumn];
                var fieldName = Utilities.FormatFieldName(field.Name);
                headerCell.Value2 = fieldName;
                var rowRange = ws.Range[ws.Cells[rowNumber, StartColumn], ws.Cells[rowNumber, EndColumn]];
                Formatting.VerticalTableRow(rowRange, rowNumber);

                rowNumber++;
            }
            
            var submitButtonRow = rowNumber + 2;
            Button submitButton = Utilities.CreateButton("submit", ws, ws.Cells[submitButtonRow, StartColumn], "Submit");
            submitButton.Click += SubmitButtonOnClick;
            ErrorData = ws.Cells[submitButtonRow, StartColumn + 2];
            ErrorData.WrapText = false;
            ErrorData.RowHeight = 14.3;

            EndRow = submitButtonRow + 1;

            titleCell.Columns.AutoFit();
        }

        public void ClearErrorData()
        {
            Utilities.ClearErrorData(ErrorData);
            ErrorData.Columns.AutoFit();
        }

        private void SubmitButtonOnClick(object sender, EventArgs e)
        {
            ClearErrorData();
            if (typeof(TRequestMessage) == typeof(ConnectPeerRequest))
            {
                var fullAddress = (string) Ws.Cells[_dataStartRow, EndColumn].Value2;
                if (fullAddress == null) return;
                var addressParts = fullAddress.Split('@');

                string pubkey;
                string host;
                switch (addressParts.Length)
                {
                    case 0:
                        return;
                    case 2:
                        pubkey = addressParts[0];
                        host = addressParts[1];
                        break;
                    default:
                        Utilities.DisplayError(ErrorData, "Error", "Invalid address, must be pubkey@ip:host");
                        return;
                }

                var permanent = Ws.Cells[_dataStartRow + 1, EndColumn].Value2;
                bool perm = permanent == null || (bool) permanent;

                var address = new LightningAddress { Host = host, Pubkey = pubkey };
                var request = new ConnectPeerRequest { Addr = address, Perm = perm };
                try
                {
                    _lApp.LndClient.ConnectPeer(request);
                    _lApp.Refresh(SheetNames.Peers);
                    Ws.Cells[_dataStartRow, EndColumn].Value2 = "";
                    Ws.Cells[_dataStartRow+1, EndColumn].Value2 = "";
                }
                catch (RpcException rpcException)
                {
                    DisplayError(rpcException);
                }
            }
            else
            {
                var request = new TRequestMessage();
                var rowNumber = _dataStartRow;
                foreach (var field in Fields)
                {
                    Range dataCell = Ws.Cells[rowNumber, EndColumn];
                    var value = dataCell.Value2;
                    if (!string.IsNullOrWhiteSpace(value?.ToString()))
                    {
                        field.Accessor.SetValue(request, dataCell.Value2);
                    }
                }

                _query(request);
            }
        }

        public void DisplayError(RpcException e)
        {
            Utilities.DisplayError(ErrorData, "Error", e);
            ErrorData.Columns.AutoFit();
        }
    }
}
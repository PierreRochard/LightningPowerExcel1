using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class ChannelsSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

    
        private Dictionary<ulong, Channel> _data;
        private int _startRow = 2;
        private int _startColumn = 2;
        private int _endColumn; 

        public ChannelsSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            this.Ws = ws;
            this.LApp = lApp;
            this._endColumn = _startColumn + Channel.Descriptor.Fields.InFieldNumberOrder().Count - 1;
            this._data = new Dictionary<ulong, Channel>();

            Tables.SetupTable<Channel>(Ws, "Channels", Channel.Descriptor);
        }
       
        public void Update(RepeatedField<Channel> data)
        {
            foreach (var channel in data)
            {
                var isCached = _data.TryGetValue(channel.ChanId, out var cachedChannel);
                if (isCached && cachedChannel.Equals(channel))
                {
                    continue;
                }

                if (isCached && !cachedChannel.Equals(channel))
                {
                    UpdateChannel(channel, cachedChannel);
                    _data[channel.ChanId] = channel;
                }
                else if (!isCached)
                {
                    AppendChannel(channel);
                    _data[channel.ChanId] = channel;
                }
            }

            foreach (var chanId in _data.Keys)
            {
                var result = data.Where(channel => channel.ChanId == chanId).ToList();
                if (result.Count == 0)
                {
                    RemoveChannel(chanId);
                }
            }

            Ws.Range["A:AZ"].Columns.AutoFit();
            Ws.Range["A:AZ"].Rows.AutoFit();
            Tables.RemoveLoadingMark(Ws);
        }

        private void RemoveChannel(ulong chanId)
        {
            var channelRow = GetChannelRow(chanId);
            Range channelRange = Ws.Range[Ws.Cells[channelRow, _startColumn], Ws.Cells[channelRow, _endColumn]];
            channelRange.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        private void AppendChannel(Channel channel)
        {
            var lastRow = GetLastRow();
            var fields = Channel.Descriptor.Fields.InDeclarationOrder();
            for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var field = fields[fieldIndex];
                var dataCell = Ws.Cells[lastRow, _startColumn + fieldIndex];
                string value = "";

                if (field.IsRepeated && field.FieldType != FieldType.Message)
                {
                    var items = (RepeatedField<string>)field.Accessor.GetValue(channel);
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
                    value = field.Accessor.GetValue(channel).ToString();
                }

                dataCell.Value2 = value;
            }
            Formatting.TableDataRow(Ws.Range[Ws.Cells[lastRow, _startColumn], Ws.Cells[lastRow, _endColumn]], lastRow % 2 == 0);
        }

        public void UpdateChannel(Channel newChannel, Channel oldChannel)
        {
            var channelRow = GetChannelRow(newChannel.ChanId);
            var fields = Channel.Descriptor.Fields.InDeclarationOrder();
            for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var field = fields[fieldIndex];
                var newValue = field.Accessor.GetValue(newChannel).ToString();
                var oldValue = field.Accessor.GetValue(oldChannel).ToString();
                if (oldValue != newValue)
                {
                    var dataCell = Ws.Cells[channelRow, _startColumn + fieldIndex];
                    string value = "";

                    if (field.IsRepeated && field.FieldType != FieldType.Message)
                    {
                        var items = (RepeatedField<string>)field.Accessor.GetValue(newChannel);
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
                        value = newValue;
                    }

                    dataCell.Value2 = value;
                }
            }

        }

        private int GetChannelRow(ulong channelChanId)
        {
            var idColumn = _startColumn;
            Range idColumnNameCell = Ws.Cells[_startRow, idColumn];
            while (idColumnNameCell.Value2 != "Chan Id")
            {
                idColumn++;
                idColumnNameCell = Ws.Cells[_startRow, idColumn];
            }

            var channelRow = _startRow;
            Range channelIdCell = Ws.Cells[channelRow, idColumn];
            while (channelIdCell.Value2.ToString() != channelChanId.ToString())
            {
                channelRow++;
                channelIdCell = Ws.Cells[channelRow, idColumn];
            }

            return channelRow;
        }

        private int GetLastRow()
        {
            var lastRow = _startRow;
            // Skip header
            lastRow++;

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

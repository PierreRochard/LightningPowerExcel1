using System;
using System.Collections.Generic;
using System.Linq;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using static Lnrpc.PendingChannelsResponse.Types;

namespace LNDExcel
{
    public class ChannelsSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        public TableSheet<PendingOpenChannel> PendingOpenChannelsTable;
        public TableSheet<Channel> OpenChannelsTable;
        public TableSheet<ClosedChannel> PendingClosingChannelsTable;
        public TableSheet<ForceClosedChannel> PendingForceClosingChannelsTable;
        public TableSheet<WaitingCloseChannel> WaitingCloseChannelsTable;
        public TableSheet<ChannelCloseSummary> ClosedChannelsTable;

        public int StartColumn = 2;
        public int StartRow = 2;
        public int EndColumn;
        public int EndRow;

        public ChannelsSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;

            PendingOpenChannelsTable = new TableSheet<PendingOpenChannel>(Ws, LApp, PendingOpenChannel.Descriptor, "channel_point", true);
            PendingOpenChannelsTable.SetupTable("Pending open", 5, StartRow);

            OpenChannelsTable = new TableSheet<Channel>(Ws, LApp, Channel.Descriptor, "chan_id");
            OpenChannelsTable.SetupTable("Open", 10, PendingOpenChannelsTable.EndRow + 2);

            PendingClosingChannelsTable = new TableSheet<ClosedChannel>(Ws, LApp, ClosedChannel.Descriptor, "channel_point", true);
            PendingClosingChannelsTable.SetupTable("Pending closing", 5, OpenChannelsTable.EndRow+2, StartColumn);

            PendingForceClosingChannelsTable = new TableSheet<ForceClosedChannel>(Ws, LApp, ForceClosedChannel.Descriptor, "channel_point", true);
            PendingForceClosingChannelsTable.SetupTable("Pending force closing", 5, PendingClosingChannelsTable.EndRow+2, StartColumn);

            WaitingCloseChannelsTable = new TableSheet<WaitingCloseChannel>(Ws, LApp, WaitingCloseChannel.Descriptor, "channel_point", true);
            WaitingCloseChannelsTable.SetupTable("Waiting for closing transaction to confirm", 5, PendingForceClosingChannelsTable.EndRow + 2, StartColumn);

            ClosedChannelsTable = new TableSheet<ChannelCloseSummary>(Ws, LApp, ChannelCloseSummary.Descriptor, "chan_id");
            ClosedChannelsTable.SetupTable("Closed", 5, WaitingCloseChannelsTable.EndRow + 2);
            
            EndRow = WaitingCloseChannelsTable.EndRow;
            EndColumn = new List<int>
            {
                PendingOpenChannelsTable.EndColumn,
                PendingClosingChannelsTable.EndColumn,
                PendingForceClosingChannelsTable.EndColumn,
                WaitingCloseChannelsTable.EndColumn,
                ClosedChannelsTable.EndColumn,
                OpenChannelsTable.EndColumn
            }.Max();
        }

        public void Update(Tuple<ListChannelsResponse, PendingChannelsResponse, ClosedChannelsResponse> r)
        {
            OpenChannelsTable.Update(r.Item1.Channels);
            var pendingChannels = r.Item2;
            PendingOpenChannelsTable.Update(pendingChannels.PendingOpenChannels);
            PendingClosingChannelsTable.Update(pendingChannels.PendingClosingChannels);
            PendingForceClosingChannelsTable.Update(pendingChannels.PendingForceClosingChannels);
            WaitingCloseChannelsTable.Update(pendingChannels.WaitingCloseChannels);
            ClosedChannelsTable.Update(r.Item3.Channels);
        }
    }
}
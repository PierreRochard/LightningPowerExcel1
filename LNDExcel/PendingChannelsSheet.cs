using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using static Lnrpc.PendingChannelsResponse.Types;

namespace LNDExcel
{
    public class PendingChannelsSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        public TableSheet<PendingOpenChannel> PendingOpenChannelsSheet;
        public TableSheet<ClosedChannel> PendingClosingChannelsSheet;
        public TableSheet<ForceClosedChannel> PendingForceClosingChannelsSheet;
        public TableSheet<WaitingCloseChannel> WaitingCloseChannelsSheet;

        public int StartColumn = 2;
        public int StartRow = 2;
        public int EndColumn;
        public int EndRow;

        public PendingChannelsSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;

            var pendingOpenWideColumns = new List<string> {"remote_node_pub"};
            PendingOpenChannelsSheet = new TableSheet<PendingOpenChannel>(Ws, LApp, PendingOpenChannel.Descriptor, "channel_point", pendingOpenWideColumns, true);
            PendingOpenChannelsSheet.SetupTable("Channels pending open", 5, StartRow);

            var pendingClosingWideColumns = new List<string> {"remote_pub_key", "closing_txid"};
            PendingClosingChannelsSheet = new TableSheet<ClosedChannel>(Ws, LApp, ClosedChannel.Descriptor, "channel_point", pendingClosingWideColumns, true);
            PendingClosingChannelsSheet.SetupTable("Channels pending closing", 5, PendingOpenChannelsSheet.EndRow+2, StartColumn);

            var pendingForceClosingWideColumns = new List<string> { "remote_pub_key", "closing_txid"};
            PendingForceClosingChannelsSheet = new TableSheet<ForceClosedChannel>(Ws, LApp, ForceClosedChannel.Descriptor, "channel_point", pendingForceClosingWideColumns, true);
            PendingForceClosingChannelsSheet.SetupTable("Channels pending force closing", 5, PendingClosingChannelsSheet.EndRow+2, StartColumn);

            var waitingCloseWideColumns = new List<string> { "remote_pub_key" };
            WaitingCloseChannelsSheet = new TableSheet<WaitingCloseChannel>(Ws, LApp, WaitingCloseChannel.Descriptor, "channel_point", waitingCloseWideColumns, true);
            WaitingCloseChannelsSheet.SetupTable("Channels waiting for closing transaction to confirm", 5, PendingForceClosingChannelsSheet.EndRow + 2, StartColumn);

            EndRow = WaitingCloseChannelsSheet.EndRow;
            EndColumn = new List<int>
            {
                PendingOpenChannelsSheet.EndColumn,
                PendingClosingChannelsSheet.EndColumn,
                PendingForceClosingChannelsSheet.EndColumn,
                WaitingCloseChannelsSheet.EndColumn
            }.Max();
        }
    }
}
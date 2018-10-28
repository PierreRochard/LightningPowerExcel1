using System.Collections.Generic;
using System.Linq;
using Lnrpc;
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

            PendingOpenChannelsSheet = new TableSheet<PendingOpenChannel>(Ws, LApp, PendingOpenChannel.Descriptor, "channel_point", true);
            PendingOpenChannelsSheet.SetupTable("Channels pending open", 5, StartRow);

            PendingClosingChannelsSheet = new TableSheet<ClosedChannel>(Ws, LApp, ClosedChannel.Descriptor, "channel_point", true);
            PendingClosingChannelsSheet.SetupTable("Channels pending closing", 5, PendingOpenChannelsSheet.EndRow+2, StartColumn);

            PendingForceClosingChannelsSheet = new TableSheet<ForceClosedChannel>(Ws, LApp, ForceClosedChannel.Descriptor, "channel_point", true);
            PendingForceClosingChannelsSheet.SetupTable("Channels pending force closing", 5, PendingClosingChannelsSheet.EndRow+2, StartColumn);

            WaitingCloseChannelsSheet = new TableSheet<WaitingCloseChannel>(Ws, LApp, WaitingCloseChannel.Descriptor, "channel_point", true);
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

        public void Update(PendingChannelsResponse result)
        {
            PendingOpenChannelsSheet.Update(result.PendingOpenChannels);
            PendingClosingChannelsSheet.Update(result.PendingClosingChannels);
            PendingForceClosingChannelsSheet.Update(result.PendingForceClosingChannels);
            WaitingCloseChannelsSheet.Update(result.WaitingCloseChannels);
        }
    }
}
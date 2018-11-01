using System;
using System.Collections.Generic;
using System.Drawing;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Button = Microsoft.Office.Tools.Excel.Controls.Button;

namespace LNDExcel
{
    public class PeersSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;

        public MessageForm<ConnectPeerRequest> PeersForm;
        public TableSheet<Peer> PeersTable;

        public int StartColumn = 2;
        public int StartRow = 2;

        public PeersSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;

            PeersForm = new MessageForm<ConnectPeerRequest>(ws, LApp, ConnectPeerRequest.Descriptor, "Connect to a peer");
            PeersTable = new TableSheet<Peer>(ws, LApp, Peer.Descriptor, "pub_key");
            PeersTable.SetupTable("Peers", startRow: PeersForm.EndRow + 2);
        }




    }
}
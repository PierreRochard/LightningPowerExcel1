using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class TransactionsSheet
    {
        public AsyncLightningApp LApp;
        public Worksheet Ws;


        public MessageForm<SendCoinsRequest, SendCoinsResponse> SendCoinsForm;
        public TableSheet<Transaction> TransactionsTable;

        public int StartColumn = 2;
        public int StartRow = 2;

        public TransactionsSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            Ws = ws;
            LApp = lApp;

            SendCoinsForm = new MessageForm<SendCoinsRequest, SendCoinsResponse>(ws, LApp, LApp.LndClient.SendCoins, SendCoinsRequest.Descriptor, "Send on-chain bitcoins");

            TransactionsTable = new TableSheet<Transaction>(ws, lApp, Transaction.Descriptor, "tx_hash");
            TransactionsTable.SetupTable("Transactions", startRow:SendCoinsForm.EndRow + 2);
        }
    }
}
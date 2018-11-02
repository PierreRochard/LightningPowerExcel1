using Lnrpc;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class TransactionsSheet
    {

        public TableSheet<Transaction> TransactionsTable;

        public TransactionsSheet(Worksheet ws, AsyncLightningApp lApp)
        {
            TransactionsTable = new TableSheet<Transaction>(ws, lApp, Transaction.Descriptor, "tx_hash");
            TransactionsTable.SetupTable("Transactions");
        }
    }
}
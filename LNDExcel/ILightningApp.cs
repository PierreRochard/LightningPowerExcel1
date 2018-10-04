using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using Lnrpc;
using LNDExcel;

namespace LNDExcel
{
    internal static class SheetNames
    {
        internal const string GetInfo = "GetInfo";
    }

    public interface ILightningApp
    {
        void RefreshGetInfo();
        string NewAddress();
    }

    public class LightningApp: ILightningApp
    {

        private readonly LndClient _lndClient;
        private readonly ThisAddIn _excelAddIn;

        public LightningApp(ThisAddIn excelAddIn)
        {
            _lndClient = new LndClient();
            _excelAddIn = excelAddIn;
        }

        public void RefreshGetInfo()
        {
            _excelAddIn.MarkAsLoading(SheetNames.GetInfo, GetInfoResponse.Descriptor);

            BackgroundWorker bw = new BackgroundWorker { WorkerSupportsCancellation = true };
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }
            bw.DoWork += bw_GetInfo;
            bw.RunWorkerCompleted += bw_GetInfo_Completed;
            bw.RunWorkerAsync();
        }

        private void bw_GetInfo(object sender, DoWorkEventArgs e)
        {
            e.Result = _lndClient.GetInfo();
        }

        private void bw_GetInfo_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            var response = (GetInfoResponse)e.Result;
            _excelAddIn.PopulateVerticalTable(SheetNames.GetInfo, GetInfoResponse.Descriptor, response);
        }

        public string NewAddress()
        {
            return _lndClient.NewAddress().Address;
        }
    }
}
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using Lnrpc;
using LNDExcel;

namespace LNDExcel
{
    internal static class SheetNames
    {
        internal const string GetInfo = "Info";
        internal const string Channels = "Channels";
        internal const string Payments = "Payments";
    }

    public interface ILightningApp
    {
        void RefreshGetInfo();
        void RefreshChannels();
        string NewAddress();
    }

    public class LightningApp: ILightningApp
    {

        public readonly LndClient LndClient;
        private readonly ThisAddIn _excelAddIn;

        public LightningApp(ThisAddIn excelAddIn)
        {
            LndClient = new LndClient();
            _excelAddIn = excelAddIn;
        }

        public void RefreshGetInfo()
        {
            _excelAddIn.MarkAsLoadingVerticalTable(SheetNames.GetInfo, GetInfoResponse.Descriptor);

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
            e.Result = LndClient.GetInfo();
        }

        private void bw_GetInfo_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            var response = (GetInfoResponse)e.Result;
            _excelAddIn.PopulateVerticalTable(SheetNames.GetInfo, GetInfoResponse.Descriptor, response);
        }

        public void RefreshChannels()
        {
            _excelAddIn.MarkAsLoadingTable(SheetNames.Channels, Channel.Descriptor);

            BackgroundWorker bw = new BackgroundWorker();
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }
            bw.DoWork += BwListChannels;
            bw.RunWorkerCompleted += BwListChannelsCompleted;
            bw.RunWorkerAsync();
        }
        
        private void BwListChannels(object sender, DoWorkEventArgs e)
        {
            e.Result = LndClient.ListChannels();
        }

        private void BwListChannelsCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var response = (ListChannelsResponse)e.Result;
            _excelAddIn.PopulateTable(SheetNames.Channels, Channel.Descriptor, response.Channels);
        }

        public void RefreshPayments()
        {
            _excelAddIn.MarkAsLoadingTable(SheetNames.Payments, Payment.Descriptor);

            BackgroundWorker bw = new BackgroundWorker();
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }
            bw.DoWork += BwListPayments;
            bw.RunWorkerCompleted += BwListPaymentsCompleted;
            bw.RunWorkerAsync();
        }

        private void BwListPayments(object sender, DoWorkEventArgs e)
        {
            e.Result = LndClient.ListPayments();
        }

        private void BwListPaymentsCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var response = (ListPaymentsResponse)e.Result;
            _excelAddIn.PopulateTable(SheetNames.Payments, Payment.Descriptor, response.Payments);
        }

        public string NewAddress()
        {
            return LndClient.NewAddress().Address;
        }
    }
}
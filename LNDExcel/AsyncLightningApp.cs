using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Grpc.Core;
using Grpc.Core.Utils;
using Lnrpc;
using Channel = Lnrpc.Channel;

namespace LNDExcel
{
    public interface IAsyncLightningApp
    {
        void RefreshGetInfo();
    }

    public class AsyncLightningApp: IAsyncLightningApp
    {

        public readonly LndClient LndClient;
        private readonly ThisAddIn _excelAddIn;

        public AsyncLightningApp(ThisAddIn excelAddIn)
        {
            LndClient = new LndClient();
            _excelAddIn = excelAddIn;
        }
        
        public void Refresh(string name)
        {
            switch (name)
            {
                case SheetNames.GetInfo:
                    RefreshGetInfo();
                    break;
                case SheetNames.Channels:
                    Refresh<ListChannelsResponse, Channel>(name, Channel.Descriptor, "Channels", LndClient.ListChannels);
                    break;
                case SheetNames.Payments:
                    Refresh<ListPaymentsResponse, Payment>(name, Payment.Descriptor, "Payments", LndClient.ListPayments);
                    break;
                case SheetNames.SendPayment:
                 //   _excelAddIn.SetupPaymentRequest();
                    break;
            }
        }

        public void SendPayment(string paymentRequest)
        {
            BackgroundWorker bw = new BackgroundWorker {WorkerReportsProgress = true};
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            int timeout = 30;
            bw.DoWork += (o, args) => BwSendPayment(o, args, paymentRequest, timeout);
            bw.ProgressChanged += BwSendPaymentOnProgressChanged;
            bw.RunWorkerCompleted += BwSendPayment_Completed;
            _excelAddIn.SendPaymentSheet.MarkSendingPayment();
            bw.RunWorkerAsync();
        }

        private void BwSendPaymentOnProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            _excelAddIn.SendPaymentSheet.UpdateSendPaymentProgress(e.ProgressPercentage);
        }

        private void BwSendPayment(object sender, DoWorkEventArgs e, string paymentRequest, int timeout)
        {
            if (sender != null)
            {
                e.Result = ProgressSend(sender, paymentRequest, timeout).GetAwaiter().GetResult();
            }
        }

        private async Task<SendResponse> ProgressSend(object sender, string paymentRequest, int timeout)
        {
            Task<SendResponse> sendTask = SendPaymentAsync(sender, paymentRequest, timeout);
            int i = 0;
            while (!sendTask.IsCompleted)
            {
                await Task.Delay(1000);
                i++;
                ((BackgroundWorker)sender).ReportProgress((int)(i * 100.0 / timeout));
            }

            return await sendTask;
        }

        private async Task<SendResponse> SendPaymentAsync(object sender, string paymentRequest, int timeout)
        {
            IAsyncStreamReader<SendResponse> stream = LndClient.SendPayment(paymentRequest, timeout);

            await stream.MoveNext(CancellationToken.None);
            return stream.Current;
        }

        private void BwSendPayment_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null && e.Result != null)
            {
                var response = (SendResponse)e.Result;
                _excelAddIn.SendPaymentSheet.PopulateSendPaymentResponse(response);
            }
            else if (e.Error != null)
            {
                var response = (RpcException)e.Error;
                _excelAddIn.SendPaymentSheet.PopulateSendPaymentError(response);
            }
        }

        public void RefreshGetInfo()
        {
            Tables.MarkAsLoadingVerticalTable(_excelAddIn.Application.Sheets[SheetNames.GetInfo], GetInfoResponse.Descriptor);

            BackgroundWorker bw = new BackgroundWorker();
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
            Tables.SetupVerticalTable(_excelAddIn.Application.Sheets[SheetNames.GetInfo], "LND Info", GetInfoResponse.Descriptor, response);
        }

        public void Refresh<TResponse, TData>(string sheetName, MessageDescriptor messageDescriptor, string propertyName, Func<IMessage> query)
        {
            Tables.MarkAsLoadingTable(_excelAddIn.Application.Sheets[sheetName]);

            BackgroundWorker bw = new BackgroundWorker();
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }
            bw.DoWork += (o, args) => BwList(o, args, query);
            switch (sheetName)
            {
                case SheetNames.Channels:
                    bw.RunWorkerCompleted += BwListChannelsCompleted;
                    break;
                default:
                    bw.RunWorkerCompleted += (o, args) => BwListCompleted<TResponse, TData>(o, args, sheetName, messageDescriptor, propertyName);
                    break;
            }
            bw.RunWorkerAsync();
        }

        private void BwList(object sender, DoWorkEventArgs e, Func<IMessage> query)
        {
            e.Result = query();
        }

        private void BwListCompleted<T, T2>(object sender, RunWorkerCompletedEventArgs e, string sheetName, MessageDescriptor messageDescriptor, string propertyName)
        {
            var response = (T)e.Result;
            var data = (RepeatedField<T2>) response.GetType().GetProperty(propertyName)?.GetValue(response, null);
            Tables.SetupTable(_excelAddIn.Application.Sheets[sheetName], "", messageDescriptor, data);
        }
        
        private void BwListChannelsCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var response = (ListChannelsResponse)e.Result;
            var data = response.Channels;
            _excelAddIn.ChannelsSheet.Update(data);
        }
    }
}
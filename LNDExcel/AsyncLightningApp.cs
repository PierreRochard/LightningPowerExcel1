using System;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Grpc.Core;
using Lnrpc;
using Channel = Lnrpc.Channel;

namespace LNDExcel
{
    public class AsyncLightningApp
    {
        public readonly LndClient LndClient;
        private readonly ThisAddIn _excelAddIn;

        public AsyncLightningApp(ThisAddIn excelAddIn)
        {
            LndClient = new LndClient();
            _excelAddIn = excelAddIn;
        }

        public void Refresh(string sheetName)
        {

            var bw = new BackgroundWorker();
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            Utilities.MarkAsLoadingTable(_excelAddIn.Application.Sheets[sheetName]);
            switch (sheetName)
            {
                case SheetNames.GetInfo:
                    bw.DoWork += (o, args) => BwQuery(o, args, LndClient.GetInfo);
                    bw.RunWorkerCompleted += (o, args) => BwVerticalListCompleted(o, args, _excelAddIn.GetInfoSheet);
                    break;
                case SheetNames.OpenChannels:
                    bw.DoWork += (o, args) => BwQuery(o, args, LndClient.ListChannels);
                    bw.RunWorkerCompleted += (o, args) =>
                        BwListCompleted<Channel, ListChannelsResponse>(o, args, _excelAddIn.ChannelsSheet);
                    break;
                case SheetNames.Payments:
                    bw.DoWork += (o, args) => BwQuery(o, args, LndClient.ListPayments);
                    bw.RunWorkerCompleted += (o, args) =>
                        BwListCompleted<Payment, ListPaymentsResponse>(o, args, _excelAddIn.PaymentsSheet);
                    break;
                default:
                    Utilities.RemoveLoadingMark(_excelAddIn.Application.Sheets[sheetName]);
                    return;
            }

            bw.RunWorkerAsync();
        }

        // ReSharper disable once UnusedParameter.Local
        private static void BwQuery(object sender, DoWorkEventArgs e, Func<IMessage> query)
        {
            e.Result = query();
        }

        // ReSharper disable once UnusedParameter.Local
        private void BwVerticalListCompleted<TResponse>(object sender, RunWorkerCompletedEventArgs e, VerticalTableSheet<TResponse> tableSheet) where TResponse : IMessage
        {
            var response = (TResponse)e.Result;
            tableSheet.Update(response);
            Utilities.RemoveLoadingMark(tableSheet.Ws);
        }
        
        // ReSharper disable once UnusedParameter.Local
        private static void BwListCompleted<TMessage, TResponse>(object sender, RunWorkerCompletedEventArgs e,
            TableSheet<TMessage> tableSheet) where TMessage : IMessage where TResponse : IMessage
        {
            var response = (TResponse)e.Result;
            var fieldDescriptor = response.Descriptor.Fields.InDeclarationOrder().FirstOrDefault(f => f.IsRepeated);
            if (fieldDescriptor == null) return;

            var data = (RepeatedField<TMessage>)fieldDescriptor.Accessor.GetValue(response);
            tableSheet.Update(data);
            Utilities.RemoveLoadingMark(tableSheet.Ws);
        }

        public void SendPayment(string paymentRequest)
        {
            var bw = new BackgroundWorker {WorkerReportsProgress = true};
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            const int timeout = 30;
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
            var sendTask = SendPaymentAsync(sender, paymentRequest, timeout);
            var i = 0;
            while (!sendTask.IsCompleted)
            {
                await Task.Delay(1000);
                i++;
                ((BackgroundWorker)sender).ReportProgress((int)(i * 100.0 / timeout));
            }

            return await sendTask;
        }

        // ReSharper disable once UnusedParameter.Local
        private async Task<SendResponse> SendPaymentAsync(object sender, string paymentRequest, int timeout)
        {
            var stream = LndClient.SendPayment(paymentRequest, timeout);

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


    }
}
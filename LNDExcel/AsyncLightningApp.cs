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
using Microsoft.Office.Interop.Excel;
using Channel = Lnrpc.Channel;

namespace LNDExcel
{
    public class AsyncLightningApp
    {
        private readonly ThisAddIn _excelAddIn;
        public LndClient LndClient;

        public AsyncLightningApp(ThisAddIn excelAddIn)
        {
            _excelAddIn = excelAddIn;
            LndClient = new LndClient();
        }

        public void Connect()
        {
            if (LndClient.Config.Host == "localhost")
            {
                _excelAddIn.NodesSheet.StartLocalNode();
            }
            LndClient.TryUnlockWallet(LndClient.Config.Password);
            Refresh(SheetNames.Payments);
            Refresh(SheetNames.OpenChannels);
            Refresh(SheetNames.Balances);
            Refresh(SheetNames.Connect);
        }

        public void Refresh(string sheetName)
        {

            var bw = new BackgroundWorker();
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            Worksheet ws = _excelAddIn.Wb.Sheets[sheetName];
            Utilities.MarkAsLoadingTable(ws);
            _excelAddIn.Wb.Sheets[sheetName].Activate();
            switch (sheetName)
            {
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
                case SheetNames.NodeLog:
                    Utilities.RemoveLoadingMark(_excelAddIn.Wb.Sheets[sheetName]);
                    break;
                case SheetNames.Connect:
                    bw.DoWork += (o, args) => BwQuery(o, args, LndClient.GetInfo);
                    bw.RunWorkerCompleted += (o, args) => BwConnectCompleted(o, args, _excelAddIn.ConnectSheet);
                    break;
                case SheetNames.Balances:
                    bw.DoWork += BwBalancesQuery;
                    bw.RunWorkerCompleted += BwBalancesCompleted;
                    break;
                default:
                    Utilities.RemoveLoadingMark(_excelAddIn.Wb.Sheets[sheetName]);
                    return;
            }

            bw.RunWorkerAsync();
        }

        private void BwBalancesCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var result = (Tuple<WalletBalanceResponse, ChannelBalanceResponse>)e.Result;
            _excelAddIn.BalancesSheet.ChannelBalanceSheet.Update(result.Item2);
            _excelAddIn.BalancesSheet.WalletBalanceSheet.Update(result.Item1);
            Utilities.RemoveLoadingMark(_excelAddIn.Wb.Sheets[SheetNames.Balances]);
        }

        private void BwBalancesQuery(object sender, DoWorkEventArgs e)
        {
            var walletBalance = LndClient.WalletBalance();
            var channelBalance = LndClient.ChannelBalance();
            var result = Tuple.Create(walletBalance, channelBalance);
            e.Result = result;
        }

        // ReSharper disable once UnusedParameter.Local
        private void BwConnectCompleted(object sender, RunWorkerCompletedEventArgs e, ConnectSheet connectSheet)
        {
            try
            {
                var response = (GetInfoResponse)e.Result;
                connectSheet.GetInfoSheet.Update(response);
                connectSheet.FormatDimensions();
            }
            catch (System.Reflection.TargetInvocationException exception)
            {
                var rpcException = (RpcException) exception.InnerException;
                _excelAddIn.ConnectSheet.DisplayError("Connect error", rpcException?.Status.Detail);
                _excelAddIn.ConnectSheet.Ws.Activate();
            }
            Utilities.RemoveLoadingMark(connectSheet.Ws);
        }

        // ReSharper disable once UnusedParameter.Local
        private static void BwQuery(object sender, DoWorkEventArgs e, Func<IMessage> query)
        {
            e.Result = query();
        }
        
        // ReSharper disable once UnusedParameter.Local
        private void BwListCompleted<TMessage, TResponse>(object sender, RunWorkerCompletedEventArgs e,
            TableSheet<TMessage> tableSheet) where TMessage : IMessage where TResponse : IMessage
        {
            try
            {
                var response = (TResponse)e.Result;
                var fieldDescriptor = response.Descriptor.Fields.InDeclarationOrder().FirstOrDefault(f => f.IsRepeated);
                if (fieldDescriptor == null) return;

                var data = (RepeatedField<TMessage>)fieldDescriptor.Accessor.GetValue(response);
                tableSheet.Update(data);
            }
            catch (System.Reflection.TargetInvocationException exception)
            {
                var rpcException = (RpcException)exception.InnerException;
                _excelAddIn.ConnectSheet.DisplayError("Connect error", rpcException?.Status.Detail);
                _excelAddIn.ConnectSheet.Ws.Activate();
            }
            Utilities.RemoveLoadingMark(tableSheet.Ws);

        }

        public void SendPayment(PayReq paymentRequest, RepeatedField<Route> routes = null)
        {
            var bw = new BackgroundWorker {WorkerReportsProgress = true};
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            const int timeout = 30;
            bw.DoWork += (o, args) => BwSendPayment(o, args, paymentRequest, timeout, routes);
            bw.ProgressChanged += BwSendPaymentOnProgressChanged;
            bw.RunWorkerCompleted += BwSendPayment_Completed;
            _excelAddIn.SendPaymentSheet.MarkSendingPayment();
            bw.RunWorkerAsync();
        }

        private void BwSendPaymentOnProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            _excelAddIn.SendPaymentSheet.UpdateSendPaymentProgress(e.ProgressPercentage);
        }

        private void BwSendPayment(object sender, DoWorkEventArgs e, PayReq paymentRequest, int timeout, RepeatedField<Route> routes = null)
        {
            if (sender != null)
            {
                e.Result = ProgressSend(sender, paymentRequest, timeout, routes).GetAwaiter().GetResult();
            }
        }

        private async Task<SendResponse> ProgressSend(object sender, PayReq paymentRequest, int timeout, RepeatedField<Route> routes = null)
        {
            var sendTask = SendPaymentAsync(sender, paymentRequest, timeout, routes);
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
        private async Task<SendResponse> SendPaymentAsync(object sender, PayReq paymentRequest, int timeout, RepeatedField<Route> routes = null)
        {
            var stream = routes != null && routes.Count != 0 ? LndClient.SendToRoute(paymentRequest, routes, timeout) : LndClient.SendPayment(paymentRequest, timeout);
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


        public PayReq DecodePaymentRequest(string payReq)
        {
            return LndClient.DecodePaymentRequest(payReq);
        }

        public StopResponse StopDaemon()
        {
            return LndClient.StopDaemon();
        }

        public QueryRoutesResponse QueryRoutes(PayReq payReq)
        {
            return LndClient.QueryRoutes(payReq.Destination, payReq.NumSatoshis, maxRoutes: 3);
        }
    }
}
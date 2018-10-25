using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Protobuf;
using Grpc.Core;
using Lnrpc;
using Microsoft.Office.Interop.Excel;
using Channel = Grpc.Core.Channel;

namespace LNDExcel
{
    public class LndClient
    {

        private void TryUnlockWallet(string password)
        {
            try
            {
                // ReSharper disable once UnusedVariable
                UnlockWalletResponse response = UnlockWallet(password);
                Thread.Sleep(3000);
            }
            catch (RpcException e)
            {
                if ("unknown service lnrpc.WalletUnlocker" == e.Status.Detail)
                {
                    // Wallet is already unlocked
                }
                else
                {
                    throw;
                }
            }
        }

        private Lightning.LightningClient GetLightningClient()
        {
            return new Lightning.LightningClient(RpcChannel);
        }

        private WalletUnlocker.WalletUnlockerClient GetWalletUnlockerClient()
        {
            return new WalletUnlocker.WalletUnlockerClient(RpcChannel);
        }

        private static string LndDataPath
        {
            get
            {
                var localAppData = Environment.GetEnvironmentVariable("LocalAppData");
                string[] lndPaths = { localAppData, "Lnd" };
                var lndPath = Path.Combine(lndPaths);
                return lndPath;
            }
        }

        private async Task AsyncAuthInterceptor(AuthInterceptorContext context, Metadata metadata)
        {
            string[] macaroonPaths = { LndDataPath, "data", "chain", "bitcoin", "testnet", "admin.macaroon" };
            var macaroonPath = Path.Combine(macaroonPaths);
            var macaroonBytes = File.ReadAllBytes(macaroonPath);
            var macaroonString = BitConverter.ToString(macaroonBytes).Replace("-", "").ToLower();
            metadata.Add(new Metadata.Entry("macaroon", macaroonString));
        }

        private Channel RpcChannel
        {
            get
            {
                var callCredentials = CallCredentials.FromInterceptor(AsyncAuthInterceptor);
                var sslCredentials = GetSslCredentials(LndDataPath);
                var channelCredentials = ChannelCredentials.Create(sslCredentials, callCredentials);
                Channel channel = new Channel("localhost", 10009, channelCredentials);
                return channel;
            }
        }

        private SslCredentials GetSslCredentials(string lndDataPath)
        {
            string[] caCertPaths = {lndDataPath , "tls.cert" };
            var caCertPath = Path.Combine(caCertPaths);
            var caCert = File.ReadAllText(caCertPath);
            var ssl = new SslCredentials(caCert);
            return ssl;
        }

        public UnlockWalletResponse UnlockWallet(string password)
        {
            var pw = ByteString.CopyFrom(password, Encoding.UTF8);
            var req = new UnlockWalletRequest { WalletPassword = pw };
            var response = GetWalletUnlockerClient().UnlockWallet(req);
            return response;
        }

        public GetInfoResponse GetInfo()
        {
            GetInfoRequest request = new GetInfoRequest();
            GetInfoResponse response = GetLightningClient().GetInfo(request);
            return response;
        }

        public NewAddressResponse NewAddress(NewAddressRequest.Types.AddressType addressType = NewAddressRequest.Types.AddressType.WitnessPubkeyHash)
        {
            NewAddressRequest request = new NewAddressRequest { Type = addressType };
            NewAddressResponse response = GetLightningClient().NewAddress(request);
            return response;
        }

        public ListChannelsResponse ListChannels()
        {
            ListChannelsRequest request = new ListChannelsRequest();
            ListChannelsResponse response = GetLightningClient().ListChannels(request);
            return response;
        } 

        public IAsyncStreamReader<SendResponse> SendPayment(string paymentRequest, int timeout)
        {
            var deadline = DateTime.UtcNow.AddSeconds(timeout);
            var duplexPaymentStreaming = GetLightningClient().SendPayment(Metadata.Empty, deadline, CancellationToken.None);
            SendRequest request = new SendRequest { PaymentRequest = paymentRequest };
            duplexPaymentStreaming.RequestStream.WriteAsync(request);
            return duplexPaymentStreaming.ResponseStream;
        }

        public SendResponse SyncSendPayment(string paymentRequest)
        {
            SendRequest request = new SendRequest {PaymentRequest = paymentRequest};
            var deadline = DateTime.UtcNow.AddSeconds(30);
            SendResponse response = GetLightningClient().SendPaymentSync(request, deadline: deadline);
            return response;
        }

        public PayReq DecodePaymentRequest(string paymentRequest)
        {
            PayReqString request = new PayReqString {PayReq = paymentRequest};
            PayReq response = GetLightningClient().DecodePayReq(request);
            return response;
        }

        public ListPaymentsResponse ListPayments()
        {
            ListPaymentsRequest request = new ListPaymentsRequest();
            ListPaymentsResponse response = GetLightningClient().ListPayments(request);
            return response;
        }

        public TransactionDetails GetTransactions()
        {
            GetTransactionsRequest request = new GetTransactionsRequest();
            TransactionDetails response = GetLightningClient().GetTransactions(request);
            return response;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Protobuf;
using Grpc.Core;
using Lnrpc;
using Channel = Grpc.Core.Channel;

namespace LNDExcel
{
    public class LndClientConfiguration
    {
        private string _macaroonPath;
        private string _macaroonString;
        private string _caCertPath;
        private string _caCertString;

        public string Network = "testnet";
        public string Host = "localhost";
        public int Port = 10009;
        public string Password = "test_password";

        public static string LndDataPath
        {
            get
            {
                var localAppData = Environment.GetEnvironmentVariable("LocalAppData");
                string[] lndPaths = { localAppData, "Lnd" };
                var lndPath = Path.Combine(lndPaths);
                return lndPath;
            }
        }
        
        public string MacaroonString
        {
            get
            {
                if (_macaroonString != null) return _macaroonString;

                var macaroonBytes = File.ReadAllBytes(MacaroonPath);
                var macaroonString = BitConverter.ToString(macaroonBytes).Replace("-", "").ToLower();
                return macaroonString;
            }
            set => _macaroonString = value;
        }

        public string MacaroonPath
        {
            get
            {
                if (_macaroonPath != null) return _macaroonPath;
                string[] macaroonPaths = { LndDataPath, "data", "chain", "bitcoin", Network, "admin.macaroon" };
                var macaroonPath = Path.Combine(macaroonPaths);
                return macaroonPath;
            }
            set => _macaroonPath = value;
        }

        private SslCredentials GetSslCredentials()
        {
            var ssl = new SslCredentials(CaCertString);
            return ssl;
        }

        public string CaCertPath
        {
            get
            {
                if (_caCertPath != null) return _caCertPath;

                string[] caCertPaths = { LndDataPath, "tls.cert" };
                var caCertPath = Path.Combine(caCertPaths);
                return caCertPath;
            }
            set => _caCertPath = value;
        }

        public string CaCertString
        {
            get
            {
                if (_caCertString != null) return _caCertString;
                var caCert = File.ReadAllText(CaCertPath);
                return caCert;
            }
            set => _caCertString = value;
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public async Task AsyncAuthInterceptor(AuthInterceptorContext context, Metadata metadata)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            metadata.Add(new Metadata.Entry("macaroon", MacaroonString));
        }
        
        public Channel RpcChannel
        {
            get
            {
                var callCredentials = CallCredentials.FromInterceptor(AsyncAuthInterceptor);
                var channelCredentials = ChannelCredentials.Create(GetSslCredentials(), callCredentials);
                var channel = new Channel(Host, Port, channelCredentials);
                return channel;
            }
        }

    }

    public class LndClient
    {
        public LndClientConfiguration Config;

        public LndClient()
        {
            Config = new LndClientConfiguration();
        }

        public void TryUnlockWallet(string password)
        {
            try
            {
                // ReSharper disable once UnusedVariable
                var response = UnlockWallet(password);
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
            return new Lightning.LightningClient(Config.RpcChannel);
        }

        private WalletUnlocker.WalletUnlockerClient GetWalletUnlockerClient()
        {
            return new WalletUnlocker.WalletUnlockerClient(Config.RpcChannel);
        }

        public StopResponse StopDaemon()
        {
            var request = new StopRequest();
            var response = GetLightningClient().StopDaemon(request);
            return response;
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
            var request = new GetInfoRequest();
            var response = GetLightningClient().GetInfo(request);
            return response;
        }

        public ConnectPeerResponse ConnectPeer(ConnectPeerRequest request)
        {
            var response = GetLightningClient().ConnectPeer(request);
            return response;
        }

        public DisconnectPeerResponse DisconnectPeer(string pubkey)
        {
            var request = new DisconnectPeerRequest{PubKey = pubkey};
            var response = GetLightningClient().DisconnectPeer(request);
            return response;
        }

        public ListPeersResponse ListPeers()
        {
            var request = new ListPeersRequest();
            var response = GetLightningClient().ListPeers(request);
            return response;
        }

        public WalletBalanceResponse WalletBalance()
        {
            var request = new WalletBalanceRequest();
            var response = GetLightningClient().WalletBalance(request);
            return response;
        }

        public TransactionDetails GetTransactions()
        {
            var request = new GetTransactionsRequest();
            var response = GetLightningClient().GetTransactions(request);
            return response;
        }

        public NewAddressResponse NewAddress(NewAddressRequest.Types.AddressType addressType = NewAddressRequest.Types.AddressType.WitnessPubkeyHash)
        {
            var request = new NewAddressRequest { Type = addressType };
            var response = GetLightningClient().NewAddress(request);
            return response;
        }

        public IAsyncStreamReader<OpenStatusUpdate> OpenChannel(string nodePubkey, long localFundingAmount, long pushSat = 0, bool isPrivate = true, 
            long minHtlcMsat = 16000, int minConfs = 1, uint remoteCsvDelay=0, long satPerByte = 0, int targetConf = 0)
        {
            var request = new OpenChannelRequest
            {
                LocalFundingAmount = localFundingAmount,
                MinConfs = minConfs,
                MinHtlcMsat = minHtlcMsat,
                NodePubkey = ByteString.CopyFromUtf8(nodePubkey),
                NodePubkeyString = nodePubkey,
                Private = isPrivate,
                PushSat = pushSat
            };
            if (remoteCsvDelay > 0) request.RemoteCsvDelay = remoteCsvDelay;
            if (satPerByte > 0) request.SatPerByte = satPerByte;
            if (targetConf > 0) request.TargetConf = targetConf;
            var response = GetLightningClient().OpenChannel(request);
            return response.ResponseStream;
        }

        public PendingChannelsResponse ListPendingChannels()
        {
            var request = new PendingChannelsRequest();
            var response = GetLightningClient().PendingChannels(request);
            return response;
        }

        public ListChannelsResponse ListChannels()
        {
            var request = new ListChannelsRequest();
            var response = GetLightningClient().ListChannels(request);
            return response;
        }

        public ChannelBalanceResponse ChannelBalance()
        {
            var request = new ChannelBalanceRequest();
            var response = GetLightningClient().ChannelBalance(request);
            return response;
        }

        public CloseStatusUpdate CloseChannel(string channelPoint, bool force)
        {
            var request = new CloseChannelRequest
            {
                ChannelPoint = new ChannelPoint
                {
                    FundingTxidBytes = ByteString.CopyFromUtf8(channelPoint.Split(':')[0]),
                    OutputIndex = uint.Parse(channelPoint.Split(':')[1])
                }
            };
            var stream = GetLightningClient().CloseChannel(request).ResponseStream;
            stream.MoveNext(CancellationToken.None);
            return stream.Current;
        }

        public ClosedChannelsResponse ListClosedChannels()
        {
            var request = new ClosedChannelsRequest
            {
                Abandoned = true,
                Breach = true,
                Cooperative = true,
                FundingCanceled = true,
                LocalForce = true,
                RemoteForce = true
            };
            var response = GetLightningClient().ClosedChannels(request);
            return response;
        }

        public ListPaymentsResponse ListPayments()
        {
            var request = new ListPaymentsRequest();
            var response = GetLightningClient().ListPayments(request);
            return response;
        }

        public PayReq DecodePaymentRequest(string paymentRequest)
        {
            var request = new PayReqString {PayReq = paymentRequest};
            var response = GetLightningClient().DecodePayReq(request);
            return response;
        }

        public QueryRoutesResponse QueryRoutes(string pubkey, long amount, long maxFixedFee = 0, long maxPercentFee = 0, int finalCltvDelta = 144, int maxRoutes = 10)
        {
            var request = new QueryRoutesRequest
            {
                Amt = amount,
                FeeLimit =
                    maxFixedFee > 0 ? new FeeLimit { Fixed = maxFixedFee } : maxPercentFee > 0 ? new FeeLimit { Percent = maxPercentFee } : null,
                FinalCltvDelta = finalCltvDelta,
                NumRoutes = maxRoutes,
                PubKey = pubkey
            };
            var response = GetLightningClient().QueryRoutes(request);
            return response;
        }

        public IAsyncStreamReader<SendResponse> SendPayment(string paymentRequest, int timeout)
        {
            var deadline = DateTime.UtcNow.AddSeconds(timeout);
            var duplexPaymentStreaming = GetLightningClient().SendPayment(Metadata.Empty, deadline, CancellationToken.None);
            var request = new SendRequest {PaymentRequest = paymentRequest};
            duplexPaymentStreaming.RequestStream.WriteAsync(request);
            return duplexPaymentStreaming.ResponseStream;
        }

        public IAsyncStreamReader<SendResponse> SendToRoute(PayReq paymentRequest, List<Route> routes, int timeout)
        {
            var deadline = DateTime.UtcNow.AddSeconds(timeout);
            var duplexPaymentStreaming = GetLightningClient().SendToRoute(Metadata.Empty, deadline, CancellationToken.None);
            var request = new SendToRouteRequest
            {
                PaymentHash = ByteString.CopyFrom(paymentRequest.PaymentHash, Encoding.UTF8)
            };
            if (routes != null) request.Routes.Add(routes);
            duplexPaymentStreaming.RequestStream.WriteAsync(request);

            return duplexPaymentStreaming.ResponseStream;
        }

        public SendResponse SyncSendPayment(string paymentRequest)
        {
            var request = new SendRequest { PaymentRequest = paymentRequest };
            var deadline = DateTime.UtcNow.AddSeconds(30);
            var response = GetLightningClient().SendPaymentSync(request, deadline: deadline);
            return response;
        }
    }
}

﻿using System;
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
    internal interface ILndClient
    {
        UnlockWalletResponse UnlockWallet(string password);
        GetInfoResponse GetInfo();
    }

    public class LndClient : ILndClient
    {
        public LndClient()
        {
            var lndDataPath = LndDataPath;
            var channel = RpcChannel;

            try
            {
                UnlockWalletResponse response = UnlockWallet("test_password");
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
    }
}

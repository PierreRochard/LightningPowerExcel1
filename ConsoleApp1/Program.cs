using System;
using Grpc.Core;
using Lnrpc;
using System.Text;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {

            Environment.SetEnvironmentVariable("GRPC_TRACE", "all");
            Environment.SetEnvironmentVariable("GRPC_VERBOSITY", "DEBUG");


            var local_app_data = Environment.GetEnvironmentVariable("LocalAppData");
            string[] lnd_paths = { local_app_data, "Lnd" };
            var lnd_path = System.IO.Path.Combine(lnd_paths);

            string[] cacert_paths = { lnd_path, "tls.cert" };
            var cacert_path = System.IO.Path.Combine(cacert_paths);
            var cacert = System.IO.File.ReadAllText(@cacert_path);

            string[] macaroon_paths = { lnd_path, "admin.macaroon" };
            var macaroon_path = System.IO.Path.Combine(macaroon_paths);
            byte[] macaroon = System.IO.File.ReadAllBytes(@macaroon_path);
            string hex = Encoding.Unicode.GetString(macaroon);
       //     string hex = BitConverter.ToString(macaroon).Replace("-", string.Empty); ;

            var ssl = new SslCredentials(cacert);

            Grpc.Core.Channel channel = new Grpc.Core.Channel("localhost", 10009, ssl);

            var wu_client = new Lnrpc.WalletUnlocker.WalletUnlockerClient(channel);
            var pw = Google.Protobuf.ByteString.CopyFrom("test_password", Encoding.ASCII);
            var req = new Lnrpc.UnlockWalletRequest { WalletPassword = pw };

            Grpc.Core.Metadata metadata = new Grpc.Core.Metadata
            {
                new Grpc.Core.Metadata.Entry("macaroon", hex)
            };
            //var response = wu_client.UnlockWallet(req, metadata);

            var client = new Lnrpc.Lightning.LightningClient(channel);
            var gi_req = new Lnrpc.GetInfoRequest();
            var response = client.GetInfo(gi_req);

            System.Threading.Thread.Sleep(1000000);
        }
    }
}

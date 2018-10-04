using LNDExcel;
using Grpc.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Lnrpc;

namespace LNDExcel.Tests
{
    [TestClass()]
    public class LndClientIntegrationTests
    {
        [TestMethod()]
        public void UnlockWalletTestWrongPassword()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act and Assert
            Assert.ThrowsException<RpcException>(() => lndClient.UnlockWallet("wrong_password"));
        }

        [TestMethod()]
        public void UnlockWalletTestRightPassword()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act and Assert
            try
            {
                UnlockWalletResponse response = lndClient.UnlockWallet("test_password");
            }
            catch (RpcException e)
            {
                // Wallet is already unlocked
                Assert.AreEqual("unknown service lnrpc.WalletUnlocker", e.Status.Detail);
            }
        }

        [TestMethod()]
        public void GetInfoTest()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act
            GetInfoResponse response = lndClient.GetInfo();

            Assert.AreEqual("0.5.0-beta commit=3b2c807288b1b7f40d609533c1e96a510ac5fa6d", response.Version);
        }

        [TestMethod()]
        public void NewAddressTest()
        {
            // Arrange
            LndClient lndClient = new LndClient();

            // Act
            NewAddressResponse response = lndClient.NewAddress();

            Assert.IsNotNull(response.Address);
        }
    }
}
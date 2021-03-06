﻿using LNDExcel;
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
            LndClient lndClient = new LndClient("test_password");

            // Act and Assert
            Assert.ThrowsException<RpcException>(() => lndClient.UnlockWallet("wrong_password"));
        }

        [TestMethod()]
        public void UnlockWalletTestRightPassword()
        {
            // Arrange
            LndClient lndClient = new LndClient("test_password");

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
            LndClient lndClient = new LndClient("test_password");

            // Act
            GetInfoResponse response = lndClient.GetInfo();

            Assert.AreEqual("0.5.0-beta commit=3b2c807288b1b7f40d609533c1e96a510ac5fa6d", response.Version);
        }

        [TestMethod()]
        public void NewAddressTest()
        {
            // Arrange
            LndClient lndClient = new LndClient("test_password");

            // Act
            NewAddressResponse response = lndClient.NewAddress();

            Assert.IsNotNull(response.Address);
        }

        [TestMethod()]
        public void ListChannelsTest()
        {
            // Arrange
            LndClient lndClient = new LndClient("test_password");

            // Act
            ListChannelsResponse response = lndClient.ListChannels();

            Assert.IsNotNull(response);
        }

        [TestMethod()]
        public void ListPaymentsTest()
        {
            LndClient lndClient = new LndClient("test_password");
            var response = lndClient.ListPayments();
            Assert.IsNotNull(response);
        }

        [TestMethod()]
        public void SendPaymentTest()
        {
            LndClient lndClient = new LndClient("test_password");
            // Todo: query a testnet lapp for a payment request
            var response = lndClient.SendPayment("", 30);
            Assert.IsNotNull(response);
        }
    }
}
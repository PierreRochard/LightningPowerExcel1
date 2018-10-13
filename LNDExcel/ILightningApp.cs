using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Lnrpc;
using LNDExcel;

namespace LNDExcel
{
    public interface ILightningApp
    {
        void RefreshGetInfo();
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
            _excelAddIn.MarkSendingPayment();
            BackgroundWorker bw = new BackgroundWorker();
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }
            bw.DoWork += (o, args) => bw_SendPayment(o, args, paymentRequest);
            bw.RunWorkerCompleted += bw_SendPayment_Completed;
            bw.RunWorkerAsync();
        }

        private void bw_SendPayment(object sender, DoWorkEventArgs e, string paymentRequest)
        {
            try
            {
                e.Result = LndClient.SendPayment(paymentRequest);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                throw;
            }
        }

        private void bw_SendPayment_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            var response = (SendResponse)e.Result;
            _excelAddIn.PopulateSendPaymentResponse(response);
        }

        public void RefreshGetInfo()
        {
            _excelAddIn.MarkAsLoadingVerticalTable(SheetNames.GetInfo, GetInfoResponse.Descriptor);

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
            _excelAddIn.PopulateVerticalTable("", SheetNames.GetInfo, GetInfoResponse.Descriptor, response);
        }

        public void Refresh<TResponse, TData>(string sheetName, MessageDescriptor messageDescriptor, string propertyName, Func<IMessage> query)
        {
            _excelAddIn.MarkAsLoadingTable(sheetName, messageDescriptor);

            BackgroundWorker bw = new BackgroundWorker();
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }
            bw.DoWork += (o, args) => BwList(o, args, query);
            bw.RunWorkerCompleted += (o, args) => BwListCompleted<TResponse, TData>(o, args, sheetName, messageDescriptor, propertyName);
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
            _excelAddIn.PopulateTable("", sheetName, messageDescriptor, data);
        }
    }
}
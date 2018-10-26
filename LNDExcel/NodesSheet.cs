using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class NodeSheet
    {
        public Worksheet Ws;
        public bool isProcessOurs = false;

        public NodeSheet(Worksheet ws)
        {
            Ws = ws;
        }

        public void StartLocalNode()
        {
            var bw = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            bw.DoWork += RunLnd;
            bw.ProgressChanged += BwRunLndOnProgressChanged;
            bw.RunWorkerAsync();
        }

        private void BwRunLndOnProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            WriteToLog(e.UserState);
        }

        private void RunLnd(object sender, DoWorkEventArgs e)
        {
            if (isProcessOurs) return;

            var processes = Process.GetProcessesByName("tempfileLND");
            foreach (var t in processes)
            {
                t.Kill();
            }

            var lndProcesses = Process.GetProcessesByName("lnd");
            if (lndProcesses.Length > 0)
            {
                WriteToLog("LND is already running, not spawning a process and thus unable to redirect log output to this tab.");
                return;
            }

            const string exeName = "tempfileLND.exe";
            var path = Path.Combine(Path.GetTempPath(), exeName);
            try
            {
                File.WriteAllBytes(path, Properties.Resources.lnd);
            }
            catch (IOException exception)
            {
                return;
            }

            var nodeProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = path,
                    Arguments = "--bitcoin.active " +
                                "--bitcoin.testnet " +
                                "--autopilot.active " +
                                "--autopilot.maxchannels=10 " +
                                "--autopilot.allocation=1 " +
                                "--autopilot.minchansize=600000 " +
                                "--autopilot.private " +
                                "--bitcoin.node=neutrino " +
                                "--neutrino.connect=faucet.lightning.community " +
                                "--debuglevel=info",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };
            nodeProcess.Start();
            isProcessOurs = true;
            nodeProcess.EnableRaisingEvents = true;
            nodeProcess.OutputDataReceived += (o, args) =>
                NodeProcessOutputDataReceived(o, args, (BackgroundWorker) sender);
            nodeProcess.ErrorDataReceived += (o, args) =>
                NodeProcessOutputDataReceived(o, args, (BackgroundWorker) sender);
            nodeProcess.BeginOutputReadLine();
            nodeProcess.BeginErrorReadLine();
            nodeProcess.WaitForExit();
        }

        private void WriteToLog(object logMessage)
        {
            var line = (Range)Ws.Rows[1];
            line.Insert(XlInsertShiftDirection.xlShiftDown);
            var cell = Ws.Cells[1, 1];
            cell.Value2 = logMessage;
        }

        // ReSharper disable once UnusedParameter.Local
        private void NodeProcessOutputDataReceived(object sender, DataReceivedEventArgs e, BackgroundWorker bw)
        {
            bw.ReportProgress(0, e.Data);
        }
        
        
    }
}
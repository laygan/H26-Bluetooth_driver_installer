using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;

namespace blu_dri_install
{
    public partial class inst_form : Form
    {
        private StringBuilder error;


        public inst_form()
        {
            InitializeComponent();

            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            error = new StringBuilder();
            try
            {
                // ドライバインストール用プロセス
                ProcessStartInfo psi = new ProcessStartInfo();
                Process ps = new Process();
                ps.StartInfo = psi;

                // インストーラ取得用
                Process instl;

                // 旧ドライバアンインストール
                Console.WriteLine("0");
                string blue = search_reg("Qualcomm Atheros Bluetooth Suite");
                if (blue != null)
                {
                    //MessageBox.Show("インストール情報を探すことができませんでした。", "Bluetoothドライバアンインストール", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    string uninstallcommand = (string)Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + blue, false).GetValue("UninstallString");
                    psi.FileName = uninstallcommand.Split(' ')[0];
                    psi.Arguments = uninstallcommand.Split(' ')[1] + " /passive /norestart";
                    ps.Start();

                    instl = WaitForStart("msiexec", "Bluetooth");
                    instl.WaitForExit();
                    WaitForExit("msiexec", "Bluetooth");
                }

                psi.FileName = ".\\QualcommAtherosBluetoothSuiteX64\\Qualcomm Atheros Bluetooth Suite (64).msi";
                psi.Arguments = "/passive /norestart";
                //ps.StartInfo = psi;

                // Bluetoothドライバインストール
                ps.Start();

                // インストーラのプロセスを取得
                instl = WaitForStart("msiexec", "Bluetooth");

                // インストーラがQualcomm...になったので、このフォームを消す
                backgroundWorker1.ReportProgress(70);

                // 検出したインストーラが終わるのを待つ
                instl.WaitForExit();
                WaitForExit("msiexec", "Bluetooth");

                // インストール完了
                System.Environment.Exit(0);
            }
            catch (Exception ex)
            {
                error.Append(ex.GetType().Name);
                error.Append(" : ");
                error.AppendLine(ex.Message);
                error.AppendLine();
                error.AppendLine(ex.StackTrace);
                MessageBox.Show(error.ToString(), "Bluetoothドライバ更新", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 指定したウィンドウタイトルを持つプロセスが、開始されるまで待つ
        /// </summary>
        /// <param name="processname"></param>
        /// <param name="windowtitle"></param>
        private Process WaitForStart(string processname,string windowtitle)
        {
            Process proc = null;
            while ((proc = runcheck(processname, windowtitle)) == null)
            {
                System.Threading.Thread.Sleep(50);
            }
            return proc;
        }

        /// <summary>
        /// 指定したウィンドウタイトルを持つプロセスが、終了するまで待つ
        /// </summary>
        /// <param name="processname"></param>
        /// <param name="windowtitle"></param>
        private void WaitForExit(string processname, string windowtitle)
        {
            while (runcheck(processname, windowtitle) != null)
            {
                System.Threading.Thread.Sleep(50);
            }
        }

        private Process runcheck(string processname, string windowtitle)
        {
            // msiexecと名乗るプロセスをすべて取得
            Process[] check = Process.GetProcessesByName(processname);

            // 取得した分だけ、繰り返す
            foreach (Process target in check)
            {
                if (target.MainWindowTitle.Contains(windowtitle))
                {
                    // タイトルが ... Bluetooth ... のプロセス発見
                    // それを返す
                    return target;
                }
            }

            // 発見できず
            return null;
        }

        /*
        private Process runcheck()
        {
            // msiexecと名乗るプロセスをすべて取得
            Process[] check = Process.GetProcessesByName("msiexec");

            // 取得した分だけ、繰り返す
            foreach (Process target in check)
            {
                if (target.MainWindowTitle.Contains("Bluetooth"))
                {
                    // タイトルが ... Bluetooth ... のプロセス発見
                    // それを返す
                    return target;
                }
            }

            // 発見できず
            return null;
        }
        */
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            /*
            if (e.ProgressPercentage == 1)
            {

            }
            else
            {
                this.Visible = false;
            }
             */
            switch (e.ProgressPercentage)
            {
                case 70:
                    this.Visible = false;
                    break;
            }
        }

        /// <summary>
        /// アンインストール情報一覧から、一致したDisplayNameのキー名を返す
        /// </summary>
        /// <param name="dname"></param>
        private string search_reg(string dname)
        {
            RegistryKey list = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall", false);
            foreach (string key in list.GetSubKeyNames())
            {
                RegistryKey appkey = list.OpenSubKey(key);
                object displayname =appkey.GetValue("DisplayName");
                appkey.Close();
                if (displayname != null && displayname.ToString().Contains(dname))
                {
                    return key;
                }

            }
            return null;
        }
    }
}

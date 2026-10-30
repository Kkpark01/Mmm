using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace McmMaintenanceReport {
    public sealed class ToolLog {
        private readonly string path;
        public ToolLog(string path) { this.path = path; }
        public static string SafeMessage(string message) {
            string value = message ?? "";
            string password = Environment.GetEnvironmentVariable("MCM_TOOL_DB_PASSWORD");
            if (!String.IsNullOrEmpty(password)) value = value.Replace(password, "***");
            return Regex.Replace(value, @"(?i)(password|pwd)\s*=\s*[^;\r\n]*", "$1=***");
        }
        private static string Csv(string value) { return "\"" + (value ?? "").Replace("\"", "\"\"") + "\""; }
        public void Write(string kind, string message) {
            DateTime now = DateTime.Now;
            string line = now.ToString("yyyy/MM/dd,HH:mm:ss", CultureInfo.InvariantCulture) + "," + Csv(kind) + "," + Csv("保守状況報告データ抽出") + "," + Csv(SafeMessage(message));
            using (var writer = new StreamWriter(path, true, Encoding.GetEncoding(932))) writer.WriteLine(line);
        }
    }
    public sealed class MainForm : Form {
        private readonly Button start = new Button(), end = new Button();
        private readonly TextBox status = new TextBox(), date = new TextBox();
        private readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        private bool busy;
        public MainForm() {
            Text = "保守状況報告データ抽出"; ClientSize = new Size(590, 190);
            FormBorderStyle = FormBorderStyle.FixedSingle; MaximizeBox = false;
            Font = new Font("MS UI Gothic", 10); StartPosition = FormStartPosition.CenterScreen;
            start.Name = "開始"; start.Text = "開始"; start.SetBounds(320, 140, 110, 32);
            end.Name = "終了"; end.Text = "終了"; end.SetBounds(450, 140, 110, 32);
            date.Name = "日時"; date.ReadOnly = true; date.SetBounds(20, 20, 210, 26);
            status.Name = "コメント"; status.ReadOnly = true; status.Multiline = true;
            status.SetBounds(20, 58, 540, 65); status.Text = "「開始」を押して実行日現在のデータを取得します。";
            Controls.AddRange(new Control[] { start, end, date, status });
            timer.Interval = 1000; timer.Tick += delegate { date.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"); };
            date.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"); timer.Start();
            start.Click += Begin; end.Click += delegate {
                if (MessageBox.Show(this, "終了しますか？", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) Close();
            };
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (busy) e.Cancel = true; };
        }
        private void UpdateStatus(string text) {
            if (InvokeRequired) { BeginInvoke(new Action<string>(UpdateStatus), text); return; }
            status.ForeColor = Color.Blue; status.Text = text;
        }
        private void Begin(object sender, EventArgs args) {
            if (busy || MessageBox.Show(this, "処理を開始しますか？", "処理開始確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) != DialogResult.Yes) return;
            busy = true; start.Enabled = false; end.Enabled = false;
            var worker = new Thread(RunReport); worker.IsBackground = true; worker.SetApartmentState(ApartmentState.STA); worker.Start();
        }
        private void RunReport() {
            var log = new ToolLog(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log.csv"));
            string savedOutput = null;
            try {
                log.Write("案内", "処理開始");
                ReportSnapshot report = ReportData.Load(ToolSettings.Load(), UpdateStatus);
                string folder = AppDomain.CurrentDomain.BaseDirectory;
                string output = Path.Combine(folder, ExcelReport.OutputName(report.AsOfDate));
                bool overwrite = false;
                if (File.Exists(output)) {
                    overwrite = (bool)Invoke(new Func<bool>(delegate {
                        return MessageBox.Show(this, "同名の帳票があります。上書きしますか？\r\n" + output, "保存確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
                    }));
                    if (!overwrite) { log.Write("案内", "上書きをキャンセル。既存帳票は変更していません。"); UpdateStatus("保存をキャンセルしました。"); return; }
                }
                UpdateStatus("エクセルデータ編集中（対象店舗 " + report.Stores.Count + "件）");
                ExcelReport.Export(report, Path.Combine(folder, ExcelReport.TemplateName), output, overwrite);
                savedOutput = output;
                log.Write("案内", "保存完了: " + report.AsOfDate.ToString("yyyy/MM/dd") + " / 店舗 " + report.Stores.Count + "件 / " + output);
                UpdateStatus("処理を終了しました。\r\n" + output);
            } catch (Exception ex) {
                string message = ToolLog.SafeMessage(ex.Message);
                try { log.Write("異常", message); } catch { message += "\r\nログも保存できませんでした。実行用フォルダの書き込み権限を確認してください。"; }
                if (savedOutput != null) message = "帳票は保存済みです。\r\n" + savedOutput + "\r\n保存後の通知・ログ処理でエラーが発生しました。\r\n" + message;
                Invoke(new Action(delegate {
                    status.ForeColor = Color.Red; status.Text = message;
                    MessageBox.Show(this, message, savedOutput == null ? "処理を中断しました" : "帳票保存後の通知", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            } finally {
                Invoke(new Action(delegate { busy = false; start.Enabled = true; end.Enabled = true; }));
            }
        }
        protected override void Dispose(bool disposing) { if (disposing) timer.Dispose(); base.Dispose(disposing); }
    }
    internal static class Program {
        [STAThread] private static void Main() {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false); Application.Run(new MainForm());
        }
    }
}

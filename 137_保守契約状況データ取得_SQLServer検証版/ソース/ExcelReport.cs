using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace McmMaintenanceReport {
    public static class StoreMapping {
        // 現行Excel.FindのMatchCase=False / MatchByte=Falseに合わせる。
        private static readonly CompareInfo NameComparison = CultureInfo.GetCultureInfo("ja-JP").CompareInfo;
        private const CompareOptions NameOptions = CompareOptions.IgnoreCase | CompareOptions.IgnoreWidth;
        public static string Normalize(string name) {
            return (name ?? "").Replace("プロジェクトチーム", "").Replace("カスタマーステーション", "").Replace(" ", "").Trim();
        }
        public static Dictionary<int, StoreResult> Resolve(IList<string> labels, IList<StoreResult> stores) {
            var mapped = new Dictionary<int, StoreResult>();
            foreach (StoreResult store in stores) {
                string search = Normalize(store.Name);
                if (search.Length == 0) throw new InvalidOperationException("店舗名を原紙へ対応付けできません: " + store.Name);
                var exact = new List<int>(); var partial = new List<int>();
                for (int i = 0; i < labels.Count; i++) {
                    string label = Normalize(labels[i]);
                    if (NameComparison.Compare(label, search, NameOptions) == 0) exact.Add(i);
                    if (label.Length > 0 && NameComparison.IndexOf(label, search, NameOptions) >= 0) partial.Add(i);
                }
                List<int> hits = exact.Count > 0 ? exact : partial;
                if (hits.Count != 1)
                    throw new InvalidOperationException("店舗「" + store.Name + "」の原紙D列の一致が" + hits.Count + "件です。店舗名を確認してください。帳票は保存しません。");
                int row = hits[0];
                if (mapped.ContainsKey(row))
                    throw new InvalidOperationException("複数の店舗が原紙の同じ行に対応します: 「" + mapped[row].Name + "」「" + store.Name + "」。帳票は保存しません。");
                CheckPrecision(store.Amount); CheckPrecision(store.Contracts); CheckPrecision(store.SupportIds);
                mapped.Add(row, store);
            }
            return mapped;
        }
        private static void CheckPrecision(long value) {
            if (value < -999999999999999L || value > 999999999999999L)
                throw new InvalidOperationException("Excelの15桁を超える値があります。丸めた帳票は保存しません。");
        }
    }

    internal sealed class ComObjects : IDisposable {
        private readonly List<object> objects = new List<object>();
        public dynamic Keep(object value) {
            if (value != null && Marshal.IsComObject(value) && !objects.Contains(value)) objects.Add(value);
            return value;
        }
        public void Dispose() {
            for (int i = objects.Count - 1; i >= 0; i--)
                try { Marshal.FinalReleaseComObject(objects[i]); } catch (InvalidComObjectException) { }
            objects.Clear();
        }
    }
    public static class ExcelReport {
        public const string TemplateName = "保守契約情報報告（月別）デフォルトフォーム.xls";
        public static string OutputName(DateTime date) { return date.ToString("MM", CultureInfo.InvariantCulture) + "月保守状況.xls"; }

        public static void Export(ReportSnapshot report, string templatePath, string outputPath, bool allowOverwrite) {
            templatePath = Path.GetFullPath(templatePath); outputPath = Path.GetFullPath(outputPath);
            if (String.Equals(templatePath, outputPath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("原紙を出力先に指定できません。");
            if (!File.Exists(templatePath)) throw new FileNotFoundException("実行用フォルダに原紙がありません。", templatePath);
            if (File.Exists(outputPath) && !allowOverwrite) throw new IOException("同名の出力ファイルがあります。上書きの確認が必要です。");
            string directory = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(directory)) throw new DirectoryNotFoundException("出力フォルダがありません。");
            string temporary = Path.Combine(directory, ".mcm-report-" + Guid.NewGuid().ToString("N") + ".xls");
            bool copied = false;
            dynamic app = null, book = null;
            try {
                File.Copy(templatePath, temporary, false); copied = true;
                // 配布原紙の読み取り専用属性は複製先だけで解除する。原紙そのものは変更しない。
                File.SetAttributes(temporary, File.GetAttributes(temporary) & ~FileAttributes.ReadOnly);
                using (var com = new ComObjects()) {
                    try {
                        Type type = Type.GetTypeFromProgID("Excel.Application");
                        if (type == null) throw new InvalidOperationException("Microsoft Excelが見つかりません。");
                        app = com.Keep(Activator.CreateInstance(type));
                        app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false;
                        app.AutomationSecurity = 3; app.AskToUpdateLinks = false;
                        dynamic books = com.Keep(app.Workbooks);
                        book = com.Keep(books.Open(temporary, 0, false));
                        if ((bool)book.ReadOnly) throw new IOException("帳票の作業用ファイルを編集できません。");
                        dynamic sheets = com.Keep(book.Worksheets);
                        dynamic sheet = com.Keep(sheets.Item["デフォルトフォーム"]);
                        dynamic headers = com.Keep(sheet.Range["E3:G3"]);
                        object[,] headerValues = (object[,])headers.Value2;
                        string[] expected = { "保守契約数", "契約サポートID数", "保守金額" };
                        for (int i = 0; i < expected.Length; i++) {
                            string label = Convert.ToString(headerValues[1, i + 1], CultureInfo.InvariantCulture).Replace("\r", "").Replace("\n", "");
                            if (label != expected[i]) throw new InvalidOperationException("原紙の見出しが一致しません: " + expected[i]);
                        }
                        dynamic names = com.Keep(sheet.Range["D4:D91"]);
                        object[,] nameValues = (object[,])names.Value2;
                        var labels = new List<string>();
                        for (int i = 1; i <= 88; i++) labels.Add(Convert.ToString(nameValues[i, 1], CultureInfo.InvariantCulture));
                        var rows = StoreMapping.Resolve(labels, report.Stores);
                        object[,] values = new object[88, 3];
                        foreach (var pair in rows) {
                            values[pair.Key, 0] = (double)pair.Value.Contracts;
                            values[pair.Key, 1] = (double)pair.Value.SupportIds;
                            values[pair.Key, 2] = (double)pair.Value.Amount;
                        }
                        dynamic numbers = com.Keep(sheet.Range["E4:G91"]);
                        numbers.Value2 = values;
                        dynamic month = com.Keep(sheet.Range["B2"]);
                        month.Value2 = report.AsOfDate.Month;
                        // コピーした原紙へ保存する。既存の書式・合計式・XLS形式を維持する。
                        sheet.Calculate();
                        book.Save();
                    } finally {
                        if (book != null) { try { book.Close(false); } catch { } }
                        if (app != null) { try { app.Quit(); } catch { } }
                    }
                }
                if (File.Exists(outputPath)) {
                    if (!allowOverwrite) throw new IOException("出力先に同名ファイルが作成されました。帳票は上書きしません。");
                    File.Replace(temporary, outputPath, null);
                } else File.Move(temporary, outputPath);
                copied = false;
            } finally {
                // この処理で作成した一時ファイルだけを削除する。既存帳票・原紙は削除しない。
                if (copied && File.Exists(temporary)) File.Delete(temporary);
            }
        }
    }
}

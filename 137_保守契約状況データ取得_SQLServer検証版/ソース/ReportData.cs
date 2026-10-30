using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace McmMaintenanceReport {
    public sealed class ToolSettings {
        public string Server, Database, Authentication;
        public bool TrustServerCertificate;
        public int ConnectionTimeout = 15, CommandTimeout = 120;

        public static ToolSettings Load() {
            var result = new ToolSettings();
            result.Server = ConfigurationManager.AppSettings["Server"];
            result.Database = ConfigurationManager.AppSettings["Database"];
            result.Authentication = ConfigurationManager.AppSettings["Authentication"];
            bool trust;
            if (!Boolean.TryParse(ConfigurationManager.AppSettings["TrustServerCertificate"], out trust))
                throw new InvalidOperationException("TrustServerCertificateにtrueまたはfalseを設定してください。");
            result.TrustServerCertificate = trust;
            result.ConnectionTimeout = ReadTimeout("ConnectionTimeoutSeconds", 15);
            result.CommandTimeout = ReadTimeout("CommandTimeoutSeconds", 120);
            return result;
        }
        private static int ReadTimeout(string key, int fallback) {
            string raw = ConfigurationManager.AppSettings[key];
            if (String.IsNullOrWhiteSpace(raw)) return fallback;
            int value;
            if (!Int32.TryParse(raw, out value) || value < 1 || value > 3600)
                throw new InvalidOperationException(key + "には1～3600秒を設定してください。");
            return value;
        }
        public string ConnectionString() {
            if (String.IsNullOrWhiteSpace(Server) || String.IsNullOrWhiteSpace(Database))
                throw new InvalidOperationException("設定ファイルのServerとDatabaseを設定してください。");
            var b = new SqlConnectionStringBuilder();
            b.DataSource = Server.Trim(); b.InitialCatalog = Database.Trim();
            b.ApplicationName = "MCM Maintenance Status Report";
            b.ConnectTimeout = ConnectionTimeout; b.Encrypt = true;
            b.TrustServerCertificate = TrustServerCertificate;
            if (String.Equals(Authentication, "Windows", StringComparison.OrdinalIgnoreCase)) b.IntegratedSecurity = true;
            else if (String.Equals(Authentication, "SqlLogin", StringComparison.OrdinalIgnoreCase)) {
                b.UserID = Environment.GetEnvironmentVariable("MCM_TOOL_DB_USER");
                b.Password = Environment.GetEnvironmentVariable("MCM_TOOL_DB_PASSWORD");
                if (String.IsNullOrEmpty(b.UserID) || String.IsNullOrEmpty(b.Password))
                    throw new InvalidOperationException("SQL認証用の環境変数MCM_TOOL_DB_USER / MCM_TOOL_DB_PASSWORDを設定し、ツールを開き直してください。");
            } else throw new InvalidOperationException("AuthenticationにWindowsまたはSqlLoginを設定してください。");
            return b.ConnectionString;
        }
    }

    public sealed class StoreResult {
        public readonly string Name;
        public readonly long Amount, Contracts, SupportIds;
        public StoreResult(string name, long amount, long contracts, long supportIds) {
            if (String.IsNullOrWhiteSpace(name)) throw new InvalidOperationException("店舗マスタの店舗名が空欄です。");
            Name = name; Amount = amount; Contracts = contracts; SupportIds = supportIds;
        }
    }
    public sealed class ReportSnapshot {
        public readonly DateTime AsOfDate;
        public readonly IList<StoreResult> Stores;
        public ReportSnapshot(DateTime date, IEnumerable<StoreResult> stores) {
            AsOfDate = date.Date;
            Stores = new List<StoreResult>(stores).AsReadOnly();
        }
    }
    public static class ReportData {
        public static ReportSnapshot Load(ToolSettings settings, Action<string> progress) {
            progress("新MCMに接続しています。");
            using (var connection = new SqlConnection(settings.ConnectionString())) {
                connection.Open();
                progress("データ取得処理中（店舗別の契約数・サポートID数・保守金額）");
                return Read(connection, settings.CommandTimeout, null);
            }
        }
        // 基準日の指定はローカル検証からのみ使用する。通常画面ではDBサーバーの当日を採用する。
        public static ReportSnapshot Read(SqlConnection connection, int timeout, DateTime? testDate) {
            using (var command = connection.CreateCommand()) {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("McmMaintenanceReport.Report.sql")) {
                    if (stream == null) throw new InvalidOperationException("内蔵SQLを読み込めません。");
                    using (var reader = new StreamReader(stream)) command.CommandText = reader.ReadToEnd();
                }
                command.CommandTimeout = timeout;
                command.Parameters.Add("@TestDate", SqlDbType.Date).Value = testDate.HasValue ? (object)testDate.Value.Date : DBNull.Value;
                using (var reader = command.ExecuteReader()) {
                    if (!reader.Read()) throw new InvalidOperationException("基準日を取得できません。");
                    DateTime asOfDate = reader.GetDateTime(0);
                    if (!reader.NextResult()) throw new InvalidOperationException("保守金額の取得結果がありません。");
                    var money = ReadValues(reader, "保守金額");
                    if (!reader.NextResult()) throw new InvalidOperationException("保守契約数の取得結果がありません。");
                    var contracts = ReadValues(reader, "保守契約数");
                    if (!reader.NextResult()) throw new InvalidOperationException("サポートID数の取得結果がありません。");
                    var supports = ReadValues(reader, "契約サポートID数");
                    if (money.Count != contracts.Count || money.Count != supports.Count)
                        throw new InvalidOperationException("店舗別集計の件数が一致しません。帳票は保存しません。");
                    var stores = new List<StoreResult>();
                    foreach (var pair in money) {
                        if (!contracts.ContainsKey(pair.Key) || !supports.ContainsKey(pair.Key))
                            throw new InvalidOperationException("店舗別集計の店舗名が一致しません。帳票は保存しません。");
                        stores.Add(new StoreResult(pair.Key, pair.Value, contracts[pair.Key], supports[pair.Key]));
                    }
                    return new ReportSnapshot(asOfDate, stores);
                }
            }
        }
        private static Dictionary<string, long> ReadValues(SqlDataReader reader, string label) {
            var values = new Dictionary<string, long>(StringComparer.Ordinal);
            while (reader.Read()) {
                if (reader.IsDBNull(0) || String.IsNullOrWhiteSpace(Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture)))
                    throw new InvalidOperationException("店舗マスタの店舗名が空欄です。帳票は保存しません。");
                string name = reader.GetString(0);
                if (reader.IsDBNull(1)) throw new InvalidOperationException("店舗「" + name + "」の" + label + "が未設定です。帳票は保存しません。");
                decimal number = Convert.ToDecimal(reader.GetValue(1), CultureInfo.InvariantCulture);
                if (Decimal.Truncate(number) != number) throw new InvalidOperationException(label + "に整数以外の値があります。");
                values.Add(name, checked((long)number));
            }
            return values;
        }
    }
}

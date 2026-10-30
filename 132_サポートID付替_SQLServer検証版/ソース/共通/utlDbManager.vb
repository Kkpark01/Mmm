Imports System.Configuration
Imports System.Data.SqlClient

' SQL Server接続と、付替処理全体を囲むトランザクションを管理する。
Public Class utlDbManager
    Implements IDisposable
    Private Con As SqlConnection
    Private Trn As SqlTransaction
    Public Property UpdateBy As String
    Public Structure ParametersType
        Dim Name As String
        Dim Type As SqlDbType
        Dim Size As Integer
        Dim Value As Object
    End Structure
    Public Function Connect() As utlResult
        Dim res As New utlResult
        Try
            Dim server As String = ConfigurationManager.AppSettings("Server")
            Dim database As String = ConfigurationManager.AppSettings("Database")
            Dim auth As String = ConfigurationManager.AppSettings("Authentication")
            If String.IsNullOrWhiteSpace(server) OrElse String.IsNullOrWhiteSpace(database) Then
                Throw New InvalidOperationException("接続先のServerとDatabaseを設定してください。")
            End If
            Dim auditMode As String = ConfigurationManager.AppSettings("AuditUserMode")
            If String.Equals(auditMode, "Windows", StringComparison.OrdinalIgnoreCase) Then
                UpdateBy = Environment.UserName
            ElseIf String.Equals(auditMode, "Fixed", StringComparison.OrdinalIgnoreCase) Then
                UpdateBy = ConfigurationManager.AppSettings("AuditUser")
            Else
                Throw New InvalidOperationException("AuditUserModeにWindowsまたはFixedを指定してください。")
            End If
            If String.IsNullOrWhiteSpace(UpdateBy) OrElse UpdateBy.Length > 50 Then
                Throw New InvalidOperationException("更新者を1～50文字で設定してください。")
            End If
            Dim builder As New SqlConnectionStringBuilder()
            builder.DataSource = server.Trim()
            builder.InitialCatalog = database.Trim()
            builder.ApplicationName = "MCM Support ID Tool"
            builder.ConnectTimeout = 10
            builder.Encrypt = True
            Dim trust As Boolean = False
            Dim trustSetting As String = ConfigurationManager.AppSettings("TrustServerCertificate")
            If Not String.IsNullOrWhiteSpace(trustSetting) AndAlso Not Boolean.TryParse(trustSetting, trust) Then
                Throw New InvalidOperationException("TrustServerCertificateにはtrueまたはfalseを指定してください。")
            End If
            builder.TrustServerCertificate = trust
            If String.Equals(auth, "Windows", StringComparison.OrdinalIgnoreCase) Then
                builder.IntegratedSecurity = True
            ElseIf String.Equals(auth, "SqlLogin", StringComparison.OrdinalIgnoreCase) Then
                Dim user As String = Environment.GetEnvironmentVariable("MCM_TOOL_DB_USER")
                Dim password As String = Environment.GetEnvironmentVariable("MCM_TOOL_DB_PASSWORD")
                If String.IsNullOrEmpty(user) OrElse String.IsNullOrEmpty(password) Then
                    Throw New InvalidOperationException("SQL認証用の環境変数を設定してください。")
                End If
                builder.UserID = user
                builder.Password = password
            Else
                Throw New InvalidOperationException("AuthenticationにWindowsまたはSqlLoginを指定してください。")
            End If
            Con = New SqlConnection(builder.ConnectionString)
            Con.Open()
            res.SetSuccess()
        Catch ex As Exception
            res.SetError("DB接続に失敗しました。接続先と認証設定を確認してください。" & vbCrLf & ex.Message)
        End Try
        Return res
    End Function
    Public Sub DisConnect()
        If Trn IsNot Nothing Then
            Try
                Trn.Rollback()
            Finally
                Trn.Dispose()
                Trn = Nothing
            End Try
        End If
        If Con IsNot Nothing Then
            Con.Dispose()
            Con = Nothing
        End If
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        DisConnect()
    End Sub
    Public Sub BeginTransaction()
        If Con Is Nothing OrElse Con.State <> ConnectionState.Open Then Throw New InvalidOperationException("DBに接続されていません。")
        Trn = Con.BeginTransaction(IsolationLevel.Serializable)
    End Sub
    Private Function CreateCommand(ByVal sql As String, ByVal prm() As ParametersType) As SqlCommand
        Dim command As New SqlCommand()
        command.Connection = Con
        command.Transaction = Trn
        Dim i As Integer = 0
        For Each item As ParametersType In prm
            Dim pos As Integer = sql.IndexOf("?"c)
            If pos < 0 Then Throw New ArgumentException("SQLの引数が一致しません。")
            Dim name As String = "@p" & i.ToString(Globalization.CultureInfo.InvariantCulture)
            sql = sql.Substring(0, pos) & name & sql.Substring(pos + 1)
            Dim param As SqlParameter = command.Parameters.Add(name, item.Type)
            If item.Size > 0 Then param.Size = item.Size
            param.Value = If(item.Value, DBNull.Value)
            i += 1
        Next
        If sql.IndexOf("?"c) >= 0 Then Throw New ArgumentException("SQLの引数が一致しません。")
        command.CommandText = sql
        Return command
    End Function
    '先に全行を読み、SQL Serverで読み取り中の接続に更新を重ねない。
    Public Function GetDataTable(ByVal sql As String, ByVal prm() As ParametersType) As DataTable
        Using command As SqlCommand = CreateCommand(sql, prm)
            Using reader As SqlDataReader = command.ExecuteReader()
                Dim result As New DataTable()
                result.Load(reader)
                Return result
            End Using
        End Using
    End Function
    Public Function ExcuteSql(ByVal sql As String, ByVal prm() As ParametersType) As Integer
        Using command As SqlCommand = CreateCommand(sql, prm)
            Return command.ExecuteNonQuery()
        End Using
    End Function
    Public Sub Commit()
        Trn.Commit()
        Trn.Dispose()
        Trn = Nothing
    End Sub
    Public Sub Rollback()
        If Trn Is Nothing Then Return
        Try
            Trn.Rollback()
        Finally
            Trn.Dispose()
            Trn = Nothing
        End Try
    End Sub
End Class

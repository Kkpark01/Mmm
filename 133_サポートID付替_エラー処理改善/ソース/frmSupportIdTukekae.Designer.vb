<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSupportIdTukekae
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.txtSakiNonyuSaki = New System.Windows.Forms.TextBox()
        Me.付け替え先納入先 = New System.Windows.Forms.Label()
        Me.txtSakiSupportId = New System.Windows.Forms.TextBox()
        Me.付け替え先 = New System.Windows.Forms.Label()
        Me.txtMotoNonyuSaki = New System.Windows.Forms.TextBox()
        Me.付け替え元納入先 = New System.Windows.Forms.Label()
        Me.txtMotoSupportId = New System.Windows.Forms.TextBox()
        Me.付け替え元 = New System.Windows.Forms.Label()
        Me.注釈 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtMessage
        '
        Me.txtMessage.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtMessage.Location = New System.Drawing.Point(12, 378)
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.ReadOnly = True
        Me.txtMessage.Size = New System.Drawing.Size(624, 23)
        Me.txtMessage.TabIndex = 23
        Me.txtMessage.TabStop = False
        '
        'btnEnd
        '
        Me.btnEnd.Font = New System.Drawing.Font("MS UI Gothic", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnEnd.Location = New System.Drawing.Point(386, 235)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(198, 78)
        Me.btnEnd.TabIndex = 22
        Me.btnEnd.Text = "終了"
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("MS UI Gothic", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnStart.Location = New System.Drawing.Point(49, 235)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(198, 78)
        Me.btnStart.TabIndex = 21
        Me.btnStart.Text = "開始"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'txtSakiNonyuSaki
        '
        Me.txtSakiNonyuSaki.Location = New System.Drawing.Point(161, 130)
        Me.txtSakiNonyuSaki.Name = "txtSakiNonyuSaki"
        Me.txtSakiNonyuSaki.ReadOnly = True
        Me.txtSakiNonyuSaki.Size = New System.Drawing.Size(404, 19)
        Me.txtSakiNonyuSaki.TabIndex = 20
        Me.txtSakiNonyuSaki.TabStop = False
        '
        '付け替え先納入先
        '
        Me.付け替え先納入先.AutoSize = True
        Me.付け替え先納入先.Location = New System.Drawing.Point(43, 133)
        Me.付け替え先納入先.Name = "付け替え先納入先"
        Me.付け替え先納入先.Size = New System.Drawing.Size(96, 12)
        Me.付け替え先納入先.TabIndex = 19
        Me.付け替え先納入先.Text = "付け替え先納入先"
        '
        'txtSakiSupportId
        '
        Me.txtSakiSupportId.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtSakiSupportId.Location = New System.Drawing.Point(161, 104)
        Me.txtSakiSupportId.MaxLength = 7
        Me.txtSakiSupportId.Name = "txtSakiSupportId"
        Me.txtSakiSupportId.Size = New System.Drawing.Size(110, 19)
        Me.txtSakiSupportId.TabIndex = 18
        '
        '付け替え先
        '
        Me.付け替え先.AutoSize = True
        Me.付け替え先.Location = New System.Drawing.Point(43, 107)
        Me.付け替え先.Name = "付け替え先"
        Me.付け替え先.Size = New System.Drawing.Size(109, 12)
        Me.付け替え先.TabIndex = 17
        Me.付け替え先.Text = "付け替え先サポートID"
        '
        'txtMotoNonyuSaki
        '
        Me.txtMotoNonyuSaki.Location = New System.Drawing.Point(161, 58)
        Me.txtMotoNonyuSaki.Name = "txtMotoNonyuSaki"
        Me.txtMotoNonyuSaki.ReadOnly = True
        Me.txtMotoNonyuSaki.Size = New System.Drawing.Size(404, 19)
        Me.txtMotoNonyuSaki.TabIndex = 16
        Me.txtMotoNonyuSaki.TabStop = False
        '
        '付け替え元納入先
        '
        Me.付け替え元納入先.AutoSize = True
        Me.付け替え元納入先.Location = New System.Drawing.Point(43, 61)
        Me.付け替え元納入先.Name = "付け替え元納入先"
        Me.付け替え元納入先.Size = New System.Drawing.Size(96, 12)
        Me.付け替え元納入先.TabIndex = 15
        Me.付け替え元納入先.Text = "付け替え元納入先"
        '
        'txtMotoSupportId
        '
        Me.txtMotoSupportId.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtMotoSupportId.Location = New System.Drawing.Point(161, 32)
        Me.txtMotoSupportId.MaxLength = 7
        Me.txtMotoSupportId.Name = "txtMotoSupportId"
        Me.txtMotoSupportId.Size = New System.Drawing.Size(110, 19)
        Me.txtMotoSupportId.TabIndex = 14
        '
        '付け替え元
        '
        Me.付け替え元.AutoSize = True
        Me.付け替え元.Location = New System.Drawing.Point(43, 35)
        Me.付け替え元.Name = "付け替え元"
        Me.付け替え元.Size = New System.Drawing.Size(109, 12)
        Me.付け替え元.TabIndex = 13
        Me.付け替え元.Text = "付け替え元サポートID"
        '
        '
        '
        '
        '
        '
        '
        '注釈
        '
        Me.注釈.AutoSize = True
        Me.注釈.ForeColor = System.Drawing.Color.Red
        Me.注釈.Location = New System.Drawing.Point(159, 161)
        Me.注釈.Name = "注釈"
        Me.注釈.Size = New System.Drawing.Size(325, 12)
        Me.注釈.TabIndex = 25
        Me.注釈.Text = "※付け替え先サポートIDが既に存在する際に納入先名を表示します"
        '
        'frmSupportIdTukekae
        '
        Me.ClientSize = New System.Drawing.Size(648, 432)
        Me.Controls.Add(Me.注釈)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.txtSakiNonyuSaki)
        Me.Controls.Add(Me.付け替え先納入先)
        Me.Controls.Add(Me.txtSakiSupportId)
        Me.Controls.Add(Me.付け替え先)
        Me.Controls.Add(Me.txtMotoNonyuSaki)
        Me.Controls.Add(Me.付け替え元納入先)
        Me.Controls.Add(Me.txtMotoSupportId)
        Me.Controls.Add(Me.付け替え元)
        Me.Name = "frmSupportIdTukekae"
        Me.Text = "保守契約管理システムサポートID付け替え"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents btnEnd As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents txtSakiNonyuSaki As System.Windows.Forms.TextBox
    Friend WithEvents 付け替え先納入先 As System.Windows.Forms.Label
    Friend WithEvents txtSakiSupportId As System.Windows.Forms.TextBox
    Friend WithEvents 付け替え先 As System.Windows.Forms.Label
    Friend WithEvents txtMotoNonyuSaki As System.Windows.Forms.TextBox
    Friend WithEvents 付け替え元納入先 As System.Windows.Forms.Label
    Friend WithEvents txtMotoSupportId As System.Windows.Forms.TextBox
    Friend WithEvents 付け替え元 As System.Windows.Forms.Label
    Friend WithEvents 注釈 As System.Windows.Forms.Label
End Class

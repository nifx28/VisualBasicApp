<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ExcelForm
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ButtonRead = New System.Windows.Forms.Button()
        Me.ButtonWrite = New System.Windows.Forms.Button()
        Me.ButtonTrigger = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ButtonRead
        '
        Me.ButtonRead.Location = New System.Drawing.Point(11, 11)
        Me.ButtonRead.Name = "ButtonRead"
        Me.ButtonRead.Size = New System.Drawing.Size(90, 45)
        Me.ButtonRead.TabIndex = 0
        Me.ButtonRead.Text = "讀取"
        Me.ButtonRead.UseVisualStyleBackColor = True
        '
        'ButtonWrite
        '
        Me.ButtonWrite.Location = New System.Drawing.Point(107, 11)
        Me.ButtonWrite.Name = "ButtonWrite"
        Me.ButtonWrite.Size = New System.Drawing.Size(90, 45)
        Me.ButtonWrite.TabIndex = 1
        Me.ButtonWrite.Text = "寫入"
        Me.ButtonWrite.UseVisualStyleBackColor = True
        '
        'ButtonTrigger
        '
        Me.ButtonTrigger.Location = New System.Drawing.Point(203, 11)
        Me.ButtonTrigger.Name = "ButtonTrigger"
        Me.ButtonTrigger.Size = New System.Drawing.Size(90, 45)
        Me.ButtonTrigger.TabIndex = 2
        Me.ButtonTrigger.Text = "觸發"
        Me.ButtonTrigger.UseVisualStyleBackColor = True
        '
        'ExcelForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(304, 67)
        Me.Controls.Add(Me.ButtonTrigger)
        Me.Controls.Add(Me.ButtonWrite)
        Me.Controls.Add(Me.ButtonRead)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "ExcelForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Microsoft Excel 2016"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ButtonRead As Button
    Friend WithEvents ButtonWrite As Button
    Friend WithEvents ButtonTrigger As Button
End Class

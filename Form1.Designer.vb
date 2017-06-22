<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.AxEScanControl1 = New AxESCANOCX2Lib.AxEScanControl2()
        CType(Me.AxEScanControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxEScanControl1
        '
        Me.AxEScanControl1.Enabled = True
        Me.AxEScanControl1.Location = New System.Drawing.Point(49, 60)
        Me.AxEScanControl1.Name = "AxEScanControl1"
        Me.AxEScanControl1.OcxState = CType(resources.GetObject("AxEScanControl1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxEScanControl1.Size = New System.Drawing.Size(192, 192)
        Me.AxEScanControl1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.AxEScanControl1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.AxEScanControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents AxEScanControl1 As AxESCANOCX2Lib.AxEScanControl2

End Class

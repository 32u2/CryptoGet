<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.BtnSetApiKey = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser1.IsWebBrowserContextMenuEnabled = False
        Me.WebBrowser1.Location = New System.Drawing.Point(0, 0)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(784, 461)
        Me.WebBrowser1.TabIndex = 0
        '
        'BtnSetApiKey
        '
        Me.BtnSetApiKey.BackColor = System.Drawing.Color.LightSteelBlue
        Me.BtnSetApiKey.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.BtnSetApiKey.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnSetApiKey.ForeColor = System.Drawing.Color.Black
        Me.BtnSetApiKey.Location = New System.Drawing.Point(683, 9)
        Me.BtnSetApiKey.Name = "BtnSetApiKey"
        Me.BtnSetApiKey.Size = New System.Drawing.Size(75, 28)
        Me.BtnSetApiKey.TabIndex = 1
        Me.BtnSetApiKey.TabStop = False
        Me.BtnSetApiKey.Text = "Set API Key"
        Me.BtnSetApiKey.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 461)
        Me.Controls.Add(Me.BtnSetApiKey)
        Me.Controls.Add(Me.WebBrowser1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CryptoGet Add-In"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents BtnSetApiKey As System.Windows.Forms.Button
End Class

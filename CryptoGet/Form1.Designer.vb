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
        Me.btnSetApiKey = New System.Windows.Forms.Button()
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
        'btnSetApiKey
        '
        Me.btnSetApiKey.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnSetApiKey.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.btnSetApiKey.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSetApiKey.ForeColor = System.Drawing.Color.Black
        Me.btnSetApiKey.Location = New System.Drawing.Point(683, 9)
        Me.btnSetApiKey.Name = "btnSetApiKey"
        Me.btnSetApiKey.Size = New System.Drawing.Size(75, 28)
        Me.btnSetApiKey.TabIndex = 1
        Me.btnSetApiKey.TabStop = False
        Me.btnSetApiKey.Text = "Set API Key"
        Me.btnSetApiKey.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 461)
        Me.Controls.Add(Me.btnSetApiKey)
        Me.Controls.Add(Me.WebBrowser1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CryptoGet Add-In"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents btnSetApiKey As System.Windows.Forms.Button
End Class

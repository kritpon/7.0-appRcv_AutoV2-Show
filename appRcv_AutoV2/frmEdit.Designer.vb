<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEdit
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmbEditInsert = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(1085, 83)
        Me.Button1.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(441, 87)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "ลบข้อมูลที่ไม่ใช้"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cmbEditInsert
        '
        Me.cmbEditInsert.Location = New System.Drawing.Point(1091, 257)
        Me.cmbEditInsert.Name = "cmbEditInsert"
        Me.cmbEditInsert.Size = New System.Drawing.Size(434, 94)
        Me.cmbEditInsert.TabIndex = 1
        Me.cmbEditInsert.Text = "cmbEdit"
        Me.cmbEditInsert.UseVisualStyleBackColor = True
        '
        'frmEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1600, 865)
        Me.Controls.Add(Me.cmbEditInsert)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
        Me.Name = "frmEdit"
        Me.Text = "frmEdit"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents cmbEditInsert As Button
End Class

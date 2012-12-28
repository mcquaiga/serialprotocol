<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TestForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.miniAtButton = New System.Windows.Forms.Button
        Me.miniMaxButton = New System.Windows.Forms.Button
        Me.ECATButton = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.itemnumTextBox = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.valueTextBox = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ReadButton = New System.Windows.Forms.Button
        Me.WriteButton = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ItemsProgressBar = New System.Windows.Forms.ProgressBar
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'miniAtButton
        '
        Me.miniAtButton.Location = New System.Drawing.Point(6, 30)
        Me.miniAtButton.Name = "miniAtButton"
        Me.miniAtButton.Size = New System.Drawing.Size(121, 37)
        Me.miniAtButton.TabIndex = 0
        Me.miniAtButton.Text = "Mini-AT"
        Me.miniAtButton.UseVisualStyleBackColor = True
        '
        'miniMaxButton
        '
        Me.miniMaxButton.Location = New System.Drawing.Point(133, 30)
        Me.miniMaxButton.Name = "miniMaxButton"
        Me.miniMaxButton.Size = New System.Drawing.Size(121, 37)
        Me.miniMaxButton.TabIndex = 1
        Me.miniMaxButton.Text = "Mini-Max"
        Me.miniMaxButton.UseVisualStyleBackColor = True
        '
        'ECATButton
        '
        Me.ECATButton.Location = New System.Drawing.Point(260, 30)
        Me.ECATButton.Name = "ECATButton"
        Me.ECATButton.Size = New System.Drawing.Size(121, 37)
        Me.ECATButton.TabIndex = 2
        Me.ECATButton.Text = "EC-AT"
        Me.ECATButton.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ECATButton)
        Me.GroupBox1.Controls.Add(Me.miniMaxButton)
        Me.GroupBox1.Controls.Add(Me.miniAtButton)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 264)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(392, 76)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Connect"
        '
        'itemnumTextBox
        '
        Me.itemnumTextBox.Location = New System.Drawing.Point(182, 93)
        Me.itemnumTextBox.Name = "itemnumTextBox"
        Me.itemnumTextBox.Size = New System.Drawing.Size(104, 20)
        Me.itemnumTextBox.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(145, 361)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(121, 29)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Disconnect"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(97, 100)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Item Number:"
        '
        'valueTextBox
        '
        Me.valueTextBox.Location = New System.Drawing.Point(183, 139)
        Me.valueTextBox.Name = "valueTextBox"
        Me.valueTextBox.Size = New System.Drawing.Size(100, 20)
        Me.valueTextBox.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(116, 142)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Value"
        '
        'ReadButton
        '
        Me.ReadButton.Location = New System.Drawing.Point(119, 190)
        Me.ReadButton.Name = "ReadButton"
        Me.ReadButton.Size = New System.Drawing.Size(89, 26)
        Me.ReadButton.TabIndex = 8
        Me.ReadButton.Text = "Read"
        Me.ReadButton.UseVisualStyleBackColor = True
        '
        'WriteButton
        '
        Me.WriteButton.Location = New System.Drawing.Point(229, 190)
        Me.WriteButton.Name = "WriteButton"
        Me.WriteButton.Size = New System.Drawing.Size(88, 26)
        Me.WriteButton.TabIndex = 9
        Me.WriteButton.Text = "Write"
        Me.WriteButton.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(164, 233)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(102, 25)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "Clear Alarms"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(309, 244)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 11
        Me.Button3.Text = "TCI"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ItemsProgressBar
        '
        Me.ItemsProgressBar.Location = New System.Drawing.Point(119, 38)
        Me.ItemsProgressBar.Name = "ItemsProgressBar"
        Me.ItemsProgressBar.Size = New System.Drawing.Size(182, 14)
        Me.ItemsProgressBar.TabIndex = 12
        '
        'TestForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(412, 402)
        Me.Controls.Add(Me.ItemsProgressBar)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.WriteButton)
        Me.Controls.Add(Me.ReadButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.valueTextBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.itemnumTextBox)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "TestForm"
        Me.Text = "TestForm"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents miniAtButton As System.Windows.Forms.Button
    Friend WithEvents miniMaxButton As System.Windows.Forms.Button
    Friend WithEvents ECATButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents itemnumTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents valueTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ReadButton As System.Windows.Forms.Button
    Friend WithEvents WriteButton As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ItemsProgressBar As System.Windows.Forms.ProgressBar
End Class

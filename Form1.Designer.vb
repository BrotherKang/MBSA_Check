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
        txtSource = New TextBox()
        Label1 = New Label()
        txtTarget = New TextBox()
        Label2 = New Label()
        btnRun = New Button()
        txtKB_Installed = New TextBox()
        SuspendLayout()
        ' 
        ' txtSource
        ' 
        txtSource.BorderStyle = BorderStyle.FixedSingle
        txtSource.Location = New Point(12, 48)
        txtSource.Name = "txtSource"
        txtSource.Size = New Size(356, 23)
        txtSource.TabIndex = 0
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 30)
        Label1.Name = "Label1"
        Label1.Size = New Size(55, 15)
        Label1.TabIndex = 1
        Label1.Text = "來源路徑"
        ' 
        ' txtTarget
        ' 
        txtTarget.BorderStyle = BorderStyle.FixedSingle
        txtTarget.Location = New Point(12, 115)
        txtTarget.Name = "txtTarget"
        txtTarget.Size = New Size(356, 23)
        txtTarget.TabIndex = 2
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(15, 89)
        Label2.Name = "Label2"
        Label2.Size = New Size(55, 15)
        Label2.TabIndex = 3
        Label2.Text = "目的路徑"
        ' 
        ' btnRun
        ' 
        btnRun.Location = New Point(140, 171)
        btnRun.Name = "btnRun"
        btnRun.Size = New Size(81, 31)
        btnRun.TabIndex = 4
        btnRun.Text = "執行"
        btnRun.UseVisualStyleBackColor = True
        ' 
        ' txtKB_Installed
        ' 
        txtKB_Installed.Location = New Point(7, 208)
        txtKB_Installed.Multiline = True
        txtKB_Installed.Name = "txtKB_Installed"
        txtKB_Installed.ScrollBars = ScrollBars.Both
        txtKB_Installed.Size = New Size(364, 378)
        txtKB_Installed.TabIndex = 5
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(383, 598)
        Controls.Add(txtKB_Installed)
        Controls.Add(btnRun)
        Controls.Add(Label2)
        Controls.Add(txtTarget)
        Controls.Add(Label1)
        Controls.Add(txtSource)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        Name = "Form1"
        Text = "MBSA資料篩選"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents txtSource As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtTarget As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnRun As Button
    Friend WithEvents txtKB_Installed As TextBox
End Class

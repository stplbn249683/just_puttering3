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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
    Me.Label2 = New System.Windows.Forms.Label()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.txtNumPoints = New System.Windows.Forms.TextBox()
    Me.txtTicker = New System.Windows.Forms.TextBox()
    Me.butUpdate = New System.Windows.Forms.Button()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.txtStrategyNo = New System.Windows.Forms.TextBox()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.lblCount = New System.Windows.Forms.Label()
    Me.butExit = New System.Windows.Forms.Button()
    Me.butRunFromFile = New System.Windows.Forms.Button()
    Me.butBrowse = New System.Windows.Forms.Button()
    Me.lblInputFileName = New System.Windows.Forms.Label()
    Me.txtFolder = New System.Windows.Forms.TextBox()
    Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
    Me.lblTickerFromFile = New System.Windows.Forms.Label()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.txtWeight = New System.Windows.Forms.TextBox()
    Me.lblRepeat = New System.Windows.Forms.Label()
    Me.Label8 = New System.Windows.Forms.Label()
    Me.Label7 = New System.Windows.Forms.Label()
    Me.Label9 = New System.Windows.Forms.Label()
    Me.Label10 = New System.Windows.Forms.Label()
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    Me.Label11 = New System.Windows.Forms.Label()
    Me.txtCategory = New System.Windows.Forms.TextBox()
    Me.Label12 = New System.Windows.Forms.Label()
    Me.txtInitalCash = New System.Windows.Forms.TextBox()
    Me.Label13 = New System.Windows.Forms.Label()
    Me.txtInterestRate = New System.Windows.Forms.TextBox()
    Me.Label14 = New System.Windows.Forms.Label()
    Me.txtNumAttempts = New System.Windows.Forms.TextBox()
    Me.txtMaxIterations = New System.Windows.Forms.TextBox()
    Me.Label15 = New System.Windows.Forms.Label()
    Me.chkIncludeInterest = New System.Windows.Forms.CheckBox()
    CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(228, 197)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(83, 13)
    Me.Label2.TabIndex = 10
    Me.Label2.Text = "Number of Days"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(56, 197)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(37, 13)
    Me.Label1.TabIndex = 6
    Me.Label1.Text = "Ticker"
    '
    'txtNumPoints
    '
    Me.txtNumPoints.Location = New System.Drawing.Point(231, 217)
    Me.txtNumPoints.Name = "txtNumPoints"
    Me.txtNumPoints.Size = New System.Drawing.Size(100, 20)
    Me.txtNumPoints.TabIndex = 11
    '
    'txtTicker
    '
    Me.txtTicker.Location = New System.Drawing.Point(25, 217)
    Me.txtTicker.Name = "txtTicker"
    Me.txtTicker.Size = New System.Drawing.Size(81, 20)
    Me.txtTicker.TabIndex = 7
    '
    'butUpdate
    '
    Me.butUpdate.Location = New System.Drawing.Point(352, 214)
    Me.butUpdate.Name = "butUpdate"
    Me.butUpdate.Size = New System.Drawing.Size(75, 23)
    Me.butUpdate.TabIndex = 12
    Me.butUpdate.Text = "Update"
    Me.butUpdate.UseVisualStyleBackColor = True
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(19, 21)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(86, 13)
    Me.Label3.TabIndex = 0
    Me.Label3.Text = "Strategy Number"
    '
    'txtStrategyNo
    '
    Me.txtStrategyNo.Location = New System.Drawing.Point(22, 37)
    Me.txtStrategyNo.Name = "txtStrategyNo"
    Me.txtStrategyNo.Size = New System.Drawing.Size(52, 20)
    Me.txtStrategyNo.TabIndex = 1
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(499, 197)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(83, 13)
    Me.Label4.TabIndex = 19
    Me.Label4.Text = "Attempt Number"
    '
    'lblCount
    '
    Me.lblCount.BackColor = System.Drawing.SystemColors.Menu
    Me.lblCount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lblCount.Location = New System.Drawing.Point(502, 219)
    Me.lblCount.Name = "lblCount"
    Me.lblCount.Size = New System.Drawing.Size(73, 27)
    Me.lblCount.TabIndex = 21
    '
    'butExit
    '
    Me.butExit.Location = New System.Drawing.Point(754, 37)
    Me.butExit.Name = "butExit"
    Me.butExit.Size = New System.Drawing.Size(75, 23)
    Me.butExit.TabIndex = 23
    Me.butExit.Text = "Exit"
    Me.butExit.UseVisualStyleBackColor = True
    '
    'butRunFromFile
    '
    Me.butRunFromFile.Location = New System.Drawing.Point(25, 294)
    Me.butRunFromFile.Name = "butRunFromFile"
    Me.butRunFromFile.Size = New System.Drawing.Size(129, 23)
    Me.butRunFromFile.TabIndex = 14
    Me.butRunFromFile.Text = "Run From File"
    Me.butRunFromFile.UseVisualStyleBackColor = True
    '
    'butBrowse
    '
    Me.butBrowse.Location = New System.Drawing.Point(333, 94)
    Me.butBrowse.Name = "butBrowse"
    Me.butBrowse.Size = New System.Drawing.Size(75, 23)
    Me.butBrowse.TabIndex = 4
    Me.butBrowse.Text = "Browse"
    Me.butBrowse.UseVisualStyleBackColor = True
    '
    'lblInputFileName
    '
    Me.lblInputFileName.AutoSize = True
    Me.lblInputFileName.Location = New System.Drawing.Point(20, 107)
    Me.lblInputFileName.Name = "lblInputFileName"
    Me.lblInputFileName.Size = New System.Drawing.Size(189, 13)
    Me.lblInputFileName.TabIndex = 2
    Me.lblInputFileName.Text = "Folder containing input and output files"
    '
    'txtFolder
    '
    Me.txtFolder.Location = New System.Drawing.Point(23, 123)
    Me.txtFolder.Name = "txtFolder"
    Me.txtFolder.Size = New System.Drawing.Size(404, 20)
    Me.txtFolder.TabIndex = 3
    '
    'ErrorProvider1
    '
    Me.ErrorProvider1.ContainerControl = Me
    '
    'lblTickerFromFile
    '
    Me.lblTickerFromFile.BackColor = System.Drawing.SystemColors.Menu
    Me.lblTickerFromFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lblTickerFromFile.Location = New System.Drawing.Point(25, 360)
    Me.lblTickerFromFile.Name = "lblTickerFromFile"
    Me.lblTickerFromFile.Size = New System.Drawing.Size(100, 19)
    Me.lblTickerFromFile.TabIndex = 16
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(27, 336)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(79, 13)
    Me.Label6.TabIndex = 15
    Me.Label6.Text = "Ticker from File"
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(496, 113)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(129, 13)
    Me.Label5.TabIndex = 17
    Me.Label5.Text = "Weight for Num of Trades"
    '
    'txtWeight
    '
    Me.txtWeight.Location = New System.Drawing.Point(499, 140)
    Me.txtWeight.Name = "txtWeight"
    Me.txtWeight.Size = New System.Drawing.Size(100, 20)
    Me.txtWeight.TabIndex = 18
    '
    'lblRepeat
    '
    Me.lblRepeat.BackColor = System.Drawing.SystemColors.Menu
    Me.lblRepeat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.lblRepeat.Location = New System.Drawing.Point(610, 219)
    Me.lblRepeat.Name = "lblRepeat"
    Me.lblRepeat.Size = New System.Drawing.Size(73, 27)
    Me.lblRepeat.TabIndex = 22
    '
    'Label8
    '
    Me.Label8.AutoSize = True
    Me.Label8.Location = New System.Drawing.Point(612, 195)
    Me.Label8.Name = "Label8"
    Me.Label8.Size = New System.Drawing.Size(42, 13)
    Me.Label8.TabIndex = 20
    Me.Label8.Text = "Repeat"
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(22, 174)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(130, 13)
    Me.Label7.TabIndex = 5
    Me.Label7.Text = "Run a single ticker symbol"
    '
    'Label9
    '
    Me.Label9.AutoSize = True
    Me.Label9.Location = New System.Drawing.Point(22, 268)
    Me.Label9.Name = "Label9"
    Me.Label9.Size = New System.Drawing.Size(315, 13)
    Me.Label9.TabIndex = 13
    Me.Label9.Text = "Run a list of ticker symbols from file s<strategy number>_input.xlsx"
    '
    'Label10
    '
    Me.Label10.BackColor = System.Drawing.SystemColors.Menu
    Me.Label10.Location = New System.Drawing.Point(228, 289)
    Me.Label10.Name = "Label10"
    Me.Label10.Size = New System.Drawing.Size(250, 120)
    Me.Label10.TabIndex = 24
    Me.Label10.Text = resources.GetString("Label10.Text")
    Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'Label11
    '
    Me.Label11.AutoSize = True
    Me.Label11.Location = New System.Drawing.Point(120, 197)
    Me.Label11.Name = "Label11"
    Me.Label11.Size = New System.Drawing.Size(49, 13)
    Me.Label11.TabIndex = 8
    Me.Label11.Text = "Category"
    '
    'txtCategory
    '
    Me.txtCategory.Location = New System.Drawing.Point(123, 217)
    Me.txtCategory.Name = "txtCategory"
    Me.txtCategory.Size = New System.Drawing.Size(86, 20)
    Me.txtCategory.TabIndex = 9
    '
    'Label12
    '
    Me.Label12.AutoSize = True
    Me.Label12.Location = New System.Drawing.Point(423, 24)
    Me.Label12.Name = "Label12"
    Me.Label12.Size = New System.Drawing.Size(67, 13)
    Me.Label12.TabIndex = 25
    Me.Label12.Text = "Initial Cash $"
    '
    'txtInitalCash
    '
    Me.txtInitalCash.Location = New System.Drawing.Point(426, 40)
    Me.txtInitalCash.Name = "txtInitalCash"
    Me.txtInitalCash.Size = New System.Drawing.Size(52, 20)
    Me.txtInitalCash.TabIndex = 26
    '
    'Label13
    '
    Me.Label13.AutoSize = True
    Me.Label13.Location = New System.Drawing.Point(527, 24)
    Me.Label13.Name = "Label13"
    Me.Label13.Size = New System.Drawing.Size(79, 13)
    Me.Label13.TabIndex = 27
    Me.Label13.Text = "Interest Rate %"
    '
    'txtInterestRate
    '
    Me.txtInterestRate.Location = New System.Drawing.Point(530, 40)
    Me.txtInterestRate.Name = "txtInterestRate"
    Me.txtInterestRate.Size = New System.Drawing.Size(52, 20)
    Me.txtInterestRate.TabIndex = 28
    '
    'Label14
    '
    Me.Label14.AutoSize = True
    Me.Label14.Location = New System.Drawing.Point(141, 21)
    Me.Label14.Name = "Label14"
    Me.Label14.Size = New System.Drawing.Size(100, 13)
    Me.Label14.TabIndex = 29
    Me.Label14.Text = "Number of Attempts"
    '
    'txtNumAttempts
    '
    Me.txtNumAttempts.Location = New System.Drawing.Point(149, 37)
    Me.txtNumAttempts.Name = "txtNumAttempts"
    Me.txtNumAttempts.Size = New System.Drawing.Size(52, 20)
    Me.txtNumAttempts.TabIndex = 30
    '
    'txtMaxIterations
    '
    Me.txtMaxIterations.Location = New System.Drawing.Point(285, 40)
    Me.txtMaxIterations.Name = "txtMaxIterations"
    Me.txtMaxIterations.Size = New System.Drawing.Size(78, 20)
    Me.txtMaxIterations.TabIndex = 31
    '
    'Label15
    '
    Me.Label15.AutoSize = True
    Me.Label15.Location = New System.Drawing.Point(273, 24)
    Me.Label15.Name = "Label15"
    Me.Label15.Size = New System.Drawing.Size(108, 13)
    Me.Label15.TabIndex = 32
    Me.Label15.Text = "Max Solver Iteriations"
    '
    'chkIncludeInterest
    '
    Me.chkIncludeInterest.AutoSize = True
    Me.chkIncludeInterest.Location = New System.Drawing.Point(530, 76)
    Me.chkIncludeInterest.Name = "chkIncludeInterest"
    Me.chkIncludeInterest.Size = New System.Drawing.Size(99, 17)
    Me.chkIncludeInterest.TabIndex = 33
    Me.chkIncludeInterest.Text = "Include Interest"
    Me.chkIncludeInterest.UseVisualStyleBackColor = True
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackColor = System.Drawing.SystemColors.ActiveBorder
    Me.ClientSize = New System.Drawing.Size(893, 465)
    Me.Controls.Add(Me.chkIncludeInterest)
    Me.Controls.Add(Me.Label15)
    Me.Controls.Add(Me.txtMaxIterations)
    Me.Controls.Add(Me.Label14)
    Me.Controls.Add(Me.txtNumAttempts)
    Me.Controls.Add(Me.Label13)
    Me.Controls.Add(Me.txtInterestRate)
    Me.Controls.Add(Me.Label12)
    Me.Controls.Add(Me.txtInitalCash)
    Me.Controls.Add(Me.Label11)
    Me.Controls.Add(Me.txtCategory)
    Me.Controls.Add(Me.Label10)
    Me.Controls.Add(Me.Label9)
    Me.Controls.Add(Me.Label7)
    Me.Controls.Add(Me.lblRepeat)
    Me.Controls.Add(Me.Label8)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.txtWeight)
    Me.Controls.Add(Me.lblTickerFromFile)
    Me.Controls.Add(Me.Label6)
    Me.Controls.Add(Me.butBrowse)
    Me.Controls.Add(Me.lblInputFileName)
    Me.Controls.Add(Me.txtFolder)
    Me.Controls.Add(Me.butRunFromFile)
    Me.Controls.Add(Me.butExit)
    Me.Controls.Add(Me.lblCount)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.txtStrategyNo)
    Me.Controls.Add(Me.butUpdate)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.txtNumPoints)
    Me.Controls.Add(Me.txtTicker)
    Me.Name = "Form1"
    Me.Text = "strategy_optimize"
    CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

  Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtNumPoints As TextBox
    Friend WithEvents txtTicker As TextBox
    Friend WithEvents butUpdate As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents txtStrategyNo As TextBox
  Friend WithEvents Label4 As Label
  Friend WithEvents lblCount As Label
    Friend WithEvents butExit As Button
    Friend WithEvents butRunFromFile As Button
    Friend WithEvents butBrowse As Button
    Friend WithEvents lblInputFileName As Label
    Friend WithEvents txtFolder As TextBox
    Friend WithEvents ErrorProvider1 As ErrorProvider
    Friend WithEvents lblTickerFromFile As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents txtWeight As TextBox
    Friend WithEvents lblRepeat As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents Label11 As Label
    Friend WithEvents txtCategory As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txtInterestRate As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents txtInitalCash As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents txtNumAttempts As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents txtMaxIterations As TextBox
    Friend WithEvents chkIncludeInterest As CheckBox
End Class

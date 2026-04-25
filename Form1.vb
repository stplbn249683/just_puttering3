Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data.SqlClient
Imports Skender.Stock.Indicators
Imports System.Reflection
Imports System.Threading

Public Class Form1
  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Dim AppPath$, error1%, sFileName$
    InitializeDefaults()
    AppPath$ = Application.StartupPath
    sFileName = AppPath$ & "\connection_string.ini"
    error1 = ReadConnectionString(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)
    sFileName = AppPath$ & "\strategy_optimize.ini"
    error1 = ReadDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)

    With UserInput
      txtTicker.Text = .ticker
      txtCategory.Text = .category
      txtNumPoints.Text = .num_for_calc.ToString.Trim
      txtStrategyNo.Text = .strategy_no.ToString.Trim
      txtNumAttempts.Text = .num_of_attempts.ToString.Trim
      txtMaxIterations.Text = .max_solver_iterations.ToString.Trim
      txtFolder.Text = .folder_name
      txtInitalCash.Text = .initial_cash.ToString("0.##").Trim
      txtInterestRate.Text = .interest_rate.ToString("0.##").Trim
      If .include_interest = "True" Then
        chkIncludeInterest.Checked = True
      Else
        chkIncludeInterest.Checked = False
      End If
    End With
    lblCount.Text = ""
    txtWeight.Text = "1.0"
  End Sub

  Private Sub butUpdate_Click(sender As Object, e As EventArgs) Handles butUpdate.Click
    Dim num_for_calc%, folder_name$
    Me.Cursor = Cursors.WaitCursor
    Dim error1%, connection_string$
    lblCount.Text = ""
    connection_string = UserInput.connection_string
    folder_name = txtFolder.Text.Trim
    pb.ticker = txtTicker.Text.Trim
    pb.category = txtCategory.Text.Trim
    num_for_calc = CInt(txtNumPoints.Text)
    pb.StrategyNo = CInt(txtStrategyNo.Text)
    pb.num_of_attempts = CInt(txtNumAttempts.Text)
    pb.max_solver_iterations = CInt(txtMaxIterations.Text)
    pb.weight_for_nt = CDbl(txtWeight.Text)
    pb.initial_cash = CDbl(txtInitalCash.Text)
    pb.interest_rate = CDbl(txtInterestRate.Text)
    pb.bIncludeInterest = False
    If chkIncludeInterest.Checked Then pb.bIncludeInterest = True

    pb.results_file = folder_name & "\s" & pb.StrategyNo.ToString.Trim & "_results.csv"
    pb.transactions_file = folder_name & "\s" & pb.StrategyNo.ToString.Trim & "_transactions.csv"
    pb.summary_file = folder_name & "\s" & pb.StrategyNo.ToString.Trim & "_summary.csv"

    Dim num_var = number_of_opt_variables(pb.StrategyNo)
    If num_var <= 0 Then
      Me.Cursor = Cursors.Default
      MessageBox.Show("number_of_opt_variables <= 0")
      Exit Sub ' nothing to do
    End If
    ReDim pb.fsave(0 To 4)

    pb.bDisplayMessage = True
    pb.bSaveDetails = True
    error1 = RunStrategy(pb.StrategyNo, pb.ticker, num_for_calc, connection_string)
    Me.Cursor = Cursors.Default
    If error1 < 0 Then Exit Sub

    UserInput.ticker = pb.ticker
    UserInput.category = pb.category
    UserInput.num_for_calc = num_for_calc
    UserInput.strategy_no = pb.StrategyNo
    UserInput.num_of_attempts = pb.num_of_attempts
    UserInput.max_solver_iterations = pb.max_solver_iterations
    UserInput.folder_name = folder_name
    UserInput.initial_cash = pb.initial_cash
    UserInput.interest_rate = pb.interest_rate
    If pb.bIncludeInterest Then
      UserInput.include_interest = "True"
    Else
      UserInput.include_interest = "False"
    End If

    Dim AppPath$, sFileName$
    AppPath$ = Application.StartupPath
    sFileName = AppPath$ & "\strategy_optimize.ini"
    error1 = SaveDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error saving file " & sFileName)

    Me.Cursor = Cursors.Default
  End Sub

  Private Sub butExit_Click(sender As Object, e As EventArgs) Handles butExit.Click
    Environment.Exit(0)
  End Sub

  Private Sub butBrowse_Click(sender As Object, e As EventArgs) Handles butBrowse.Click
    Dim s1$

    'Directory.SetCurrentDirectory(CurrentDir)
    s1 = txtFolder.Text.Trim
    If s1.Length > 0 Then
      If Directory.Exists(s1) = False Then
        s1 = ""
      End If
    End If

    FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
    FolderBrowserDialog1.SelectedPath = s1

    If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
      s1 = FolderBrowserDialog1.SelectedPath
      txtFolder.Text = s1
      ErrorProvider1.SetError(txtFolder, "")
    End If
  End Sub

  Private Sub butRunFromFile_Click(sender As Object, e As EventArgs) Handles butRunFromFile.Click
    Dim folder_name$

    Dim error1%, connection_string$, limit%, num_tickers%, i%
    lblCount.Text = ""
    folder_name = txtFolder.Text.Trim
    connection_string = UserInput.connection_string
    pb.StrategyNo = CInt(txtStrategyNo.Text)
    pb.num_of_attempts = CInt(txtNumAttempts.Text)
    pb.max_solver_iterations = CInt(txtMaxIterations.Text)
    pb.weight_for_nt = CDbl(txtWeight.Text)
    pb.initial_cash = CDbl(txtInitalCash.Text)
    pb.interest_rate = CDbl(txtInterestRate.Text)
    pb.bIncludeInterest = False
    If chkIncludeInterest.Checked Then pb.bIncludeInterest = True

    Dim input_file$
    input_file = folder_name & "\s" & pb.StrategyNo.ToString.Trim & "_input.xlsx"
    pb.results_file = folder_name & "\s" & pb.StrategyNo.ToString.Trim & "_results.csv"
    pb.transactions_file = folder_name & "\s" & pb.StrategyNo.ToString.Trim & "_transactions.csv"
    pb.summary_file = folder_name & "\s" & pb.StrategyNo.ToString.Trim & "_summary.csv"

    If (Dir(input_file) = "") Or Not File.Exists(input_file) Then
      MessageBox.Show("Error finding file " & input_file)
      Exit Sub
    End If

    limit = 300
    Dim tickers_in$(), categories$(), num_points%()
    Dim dt = ReadExcelFileSAX(input_file)
    If IsNothing(dt) Then Exit Sub
    num_tickers = dt.Rows.Count
    ReDim tickers_in$(0 To num_tickers - 1), categories$(0 To num_tickers - 1), num_points%(0 To num_tickers - 1)
    For i = 0 To num_tickers - 1
      tickers_in$(i) = dt.Rows(i).Item(0)
      categories$(i) = dt.Rows(i).Item(1)
      num_points%(i) = CInt(dt.Rows(i).Item(2))
    Next

    Dim num_var = number_of_opt_variables(pb.StrategyNo)
    If num_var <= 0 Then
      Me.Cursor = Cursors.Default
      MessageBox.Show("number_of_opt_variables <= 0")
      Exit Sub ' nothing to do
    End If
    ReDim pb.fsave(0 To 4)

    pb.bDisplayMessage = False
    pb.bSaveDetails = False
    For i = 0 To num_tickers - 1
      Me.Cursor = Cursors.WaitCursor
      lblTickerFromFile.Text = tickers_in(i)
      pb.ticker = tickers_in(i)
      pb.category = categories(i)
      error1 = RunStrategy(pb.StrategyNo, tickers_in(i), num_points(i), connection_string)
      Me.Cursor = Cursors.Default
      If error1 < 0 Then Exit Sub
    Next
    MessageBox.Show("Finished")

    UserInput.strategy_no = pb.StrategyNo
    UserInput.num_of_attempts = pb.num_of_attempts
    UserInput.max_solver_iterations = pb.max_solver_iterations
    UserInput.folder_name = folder_name
    UserInput.initial_cash = pb.initial_cash
    UserInput.interest_rate = pb.interest_rate
    If pb.bIncludeInterest Then
      UserInput.include_interest = "True"
    Else
      UserInput.include_interest = "False"
    End If

    Dim AppPath$, sFileName$
    AppPath$ = Application.StartupPath
    sFileName = AppPath$ & "\strategy_optimize.ini"
    error1 = SaveDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error saving file " & sFileName)
  End Sub
End Class

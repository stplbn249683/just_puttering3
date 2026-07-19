' This program uses backtested historical stock price data from the close to compute globally optimized parameters for the stock holding
' periods that result in the largest gain for the criteria that were used.

' Modified on 18Jul26 to correct an error where, if the "Include Interest" checkbox was checked, the gain from the interest was
' included twice in the summary file columns for "% of return for hold" and "Return %".  Also, the interest was being overestimated
' because the number of days was calculated using actual days but the interest rate calculation assumed 252 market days.
' Last modified on 18Jul26

Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data.SqlClient
Imports Skender.Stock.Indicators
Imports NLoptNet
Imports Microsoft.SqlServer
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Bibliography
Imports DocumentFormat.OpenXml.Wordprocessing
Imports DocumentFormat.OpenXml.InkML
Imports DocumentFormat.OpenXml.Math
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports DocumentFormat.OpenXml.ExtendedProperties
Imports System.Globalization

Structure INPUTTYPE
  Dim connection_string$
  Dim ticker$
  Dim category$
  Dim num_for_calc%
  Dim strategy_no%
  Dim num_of_attempts%
  Dim max_solver_iterations%
  Dim folder_name$
  Dim initial_cash#
  Dim interest_rate#
  Dim include_interest$
End Structure
Public Structure CellInfo
  Public row%
  Public column%
  Public value$

  Public Sub New(ByVal _row%, ByVal _column%, ByVal _value$)
    row = _row
    column = _column
    value = _value
  End Sub
End Structure

Structure PublicAccessType
  Dim bSaveResults As Boolean, bSaveDetails As Boolean, bDisplayMessage As Boolean, bIncludeInterest As Boolean
  Dim ticker$, category$, StrategyNo%, weight_for_nt#, weight_for_nt1#, ratio#, ratio_old#, results_file$, transactions_file$, summary_file$
  Dim fsave#(), x_save#(), x_type$(), num_variables%, num_trades%, best_score$, num_trades_goal%, initial_cash#, interest_rate#
  Dim max_gain#, perc_gain_per_day_weight#, min_perc_days_in_market#, perc_days_in_market_weight#, win_rate_weight#, perc_max_drawdown_weight#, gl_ratio_weight#
  Dim min_num_trades_per_year#, perc_return_per_year_weight#, sharpe_ratio_weight#, sell_date_index%, num_of_attempts%
  Dim num_for_resize%, max_solver_iterations%, num_used%
End Structure
Module Module1
  Public UserInput As INPUTTYPE
  Public pb As PublicAccessType

  Sub optimize()
    Dim i%
    pb.num_variables = number_of_opt_variables(pb.StrategyNo)
    If pb.num_variables <= 0 Then Exit Sub ' nothing to do

    Dim initialValue#()
    ReDim initialValue#(0 To pb.num_variables - 1)
    ReDim pb.x_save(0 To pb.num_variables - 1)
    ReDim pb.x_type(0 To pb.num_variables - 1)
    For i = 0 To pb.num_variables - 1
      pb.x_type(i) = "d"
    Next

    Dim j%, error2%, fi1#, num_trades_save%, num_trades_goal_save%, num_trades_old%
    Dim bSave As Boolean
    pb.max_gain = -1.0E+20
    pb.best_score = ""
    pb.weight_for_nt1 = pb.weight_for_nt
    Dim max_saved_score = 0.0

    'Repeat with a higher weight if the number of trades is too low
    For k = 0 To 9
      Form1.lblRepeat.Text = k.ToString.Trim
      'Since each result is somewhat different, do it several times and use the one with the best score
      num_trades_old = 10000
      For j = 0 To pb.num_of_attempts - 1
        Form1.lblCount.Text = j.ToString.Trim
        System.Windows.Forms.Application.DoEvents()
        pb.num_trades = 0
        For i = 0 To 4
          pb.fsave(i) = 0.0
        Next

        Dim num_variables1 As UInteger
        num_variables1 = pb.num_variables
        Dim max_iterations As Integer
        max_iterations = pb.max_solver_iterations
        pb.bSaveResults = False
        ' I don't think that NLoptAlgorithm.LN_SBPLX is used for NLoptAlgorithm.GN_ESCH but I'm leaving it there
        ' as a placeholder
        Dim solver As New NLoptSolver(NLoptAlgorithm.GN_ESCH, num_variables1, 0.000001, max_iterations, NLoptAlgorithm.LN_SBPLX)
        Using solver
          'solver.SetRelativeToleranceOnFunctionValue(-1.0) ' disables the criteria
          'solver.SetRelativeToleranceOnOptimizationParameter(-1.0)
          set_bounds(pb.StrategyNo, solver, initialValue)

          Dim f1 As Func(Of Double(), Double) = Function(variables)
                                                  Dim error1%, fi#
                                                  error1 = RunStrategy1(variables, fi)
                                                  If error1 < 0 Then
                                                    MessageBox.Show("Error in RunStrategy1")
                                                    Environment.Exit(0)
                                                  End If
                                                  Return (fi)
                                                End Function

          solver.SetMaxObjective(f1)
          'Dim constraints As Action(Of Double(), Double()) = Function(variables, variables1)
          '
          'End Function
          'solver.AddLessOrEqualZeroConstraints(constraints)
          Dim finalScore As Double?

          Dim Result = solver.Optimize(initialValue, finalScore)
          Debug.Assert(Result = NloptResult.MAXEVAL_REACHED)

          error2 = RunStrategy1(initialValue, fi1)
          If error2 < 0 Then
            MessageBox.Show("Error in RunStrategy1")
            Environment.Exit(0)
          End If

          bSave = False
          Dim score = 0.0
          If finalScore IsNot Nothing Then score = finalScore
          pb.ratio_old = CDbl(num_trades_old) / CDbl(pb.num_trades_goal)
          pb.ratio = CDbl(pb.num_trades) / CDbl(pb.num_trades_goal)
          If (j = 0) Then
            bSave = True
          ElseIf (pb.ratio_old <= 0.875) And (pb.ratio > 0.875) Then
            bSave = True
          ElseIf ((pb.ratio >= pb.ratio_old) Or (pb.ratio > 0.875)) And score > max_saved_score Then
            bSave = True
          End If

          If bSave Then
            max_saved_score = score
            pb.max_gain = pb.fsave(0)
            For i = 0 To pb.num_variables - 1
              pb.x_save(i) = initialValue(i)
            Next
            num_trades_save = pb.num_trades
            num_trades_old = pb.num_trades
            num_trades_goal_save = pb.num_trades_goal
            pb.best_score = ""
            If finalScore IsNot Nothing Then pb.best_score = finalScore.ToString("0.0000")
          End If
        End Using
      Next
      pb.ratio = CDbl(num_trades_save) / CDbl(num_trades_goal_save)
      If pb.ratio > 0.875 Then Exit For
      If k <= 2 Then
        pb.weight_for_nt1 += 1.0
      Else
        pb.weight_for_nt1 += CDbl(k - 1)
      End If
    Next

    'fsave(0)...fsave(4) = gain_per_day,gain1,perc_days_in_market,win_rate,perc_return_per_year
    pb.bSaveResults = True
    error2 = RunStrategy1(pb.x_save, fi1)
    If error2 < 0 Then
      MessageBox.Show("Error in RunStrategy1")
      Environment.Exit(0)
    End If

    If pb.max_gain < -1.0E+19 Then
      MessageBox.Show("No solution was found that satisfied all of the conditions")
      'MessageBox.Show(x_save(0).ToString & "," & x_save(1).ToString & "," & x_save(2).ToString & "," & x_save(3).ToString & "," & x_save(4).ToString & "," & x_save(5).ToString & "," & x_save(6).ToString & "," & x_save(7).ToString & Environment.NewLine &
      'fsave(0).ToString & "," & fsave(1).ToString & "," & num_trades.ToString)
    End If
  End Sub
  Function RunStrategy(StrategyNo%, ticker$, num_for_calc%, connection_string$)
    RunStrategy = -1
    Dim error1%
    Dim max_num_points%, num_from_db%
    'Dim quotes As IEnumerable(Of Skender.Stock.Indicators.Quote)

    max_num_points = num_for_calc + 720 'add some points so that errors have time to die out

    pb1.quotes = GetQuotes(max_num_points, ticker, connection_string).Validate()
    num_from_db = pb1.quotes.Count
    If num_from_db <= 0 Then
      MessageBox.Show("ticker symbol not In database")
      Exit Function
    End If

    If num_from_db <= 10 Or num_for_calc > num_from_db Then
      MessageBox.Show("Not enough points for ticker symbol in database")
      Exit Function
    End If

    Call GetQuoteLists(pb1.quotes, pb1.date1, pb1.high, pb1.low, pb1.open, pb1.close, pb1.volume)
    error1 = calculate_indicators(StrategyNo, max_num_points, num_for_calc, pb1.quotes, connection_string$)
    If error1 < 0 Then
      MessageBox.Show("Error in calculate_indicators")
      Exit Function
    End If

    Call optimize()
    RunStrategy = 0
  End Function

  Function RunStrategy1(x#(), ByRef fi#)
    RunStrategy1 = -1

    Dim bDaySold, bDayBought As Boolean, i%, j%
    Dim total_num_shares#, gain_for_hold#
    Dim cost#, win_rate#, perc_days_in_market#, gain_per_day_hold#
    Dim price_sold#, fraction_sold#, count_gain%, count_loss%, gain_for_count#, loss_for_count#, price_if_sold#
    Dim gain1#, gain_per_day#, perc_gain_per_day#, gain_from_interest#, perc_gain_for_count#, perc_loss_for_count#
    Dim perc_return#, perc_return_hold#, perc_of_return_for_hold#, perc_return_per_year#, perc_return_per_day_hold#
    Dim day_bought_index%, days_return#, days_return_hold#
    Dim returns, adjusted_returns, returns_hold, adjusted_returns_hold, total_value, total_value_hold As New List(Of Double)

    Dim error1%
    error1 = calculate_indicators1(x)
    If error1 < 0 Then
      MessageBox.Show("Error in calculate_indicators1")
      Exit Function
    End If

    total_num_shares = 0#
    cost = 0.0
    count_gain = 0
    count_loss = 0
    gain_for_count = 0.0
    loss_for_count = 0.0
    perc_gain_for_count = 0.0
    perc_loss_for_count = 0.0
    gain_from_interest = 0.0
    perc_return = 0.0
    perc_return_hold = 0.0
    days_return = 0.0
    days_return_hold = 0.0
    Dim max_drawdown# = 0.0
    Dim max_drawdown_for_hold# = 0.0
    Dim max_loss# = 1.0E+25
    Dim max_loss_for_hold# = 1.0E+25
    Dim amount# = 0.0
    Dim amount_hold# = 0.0


    Dim sFileName1$ = "", sFileName2$ = ""
    Dim s1$ = "", s2$ = "", s3$ = ""
    Dim writer1 As StreamWriter = Nothing
    Dim writer2 As StreamWriter = Nothing
    If pb.bSaveResults AndAlso pb.bSaveDetails Then
      sFileName1 = pb.transactions_file
      If File.Exists(sFileName1) Then File.Delete(sFileName1)
      sFileName2 = pb.results_file
      If File.Exists(sFileName2) Then File.Delete(sFileName2)
    End If
    Try
      If pb.bSaveResults AndAlso pb.bSaveDetails Then
        writer1 = New StreamWriter(sFileName1)
        writer2 = New StreamWriter(sFileName2)
        s1 = "Transaction,Date,Total Shares,Current Price,Gain/Loss,Gain/Loss for Hold"
        s2 = "Date,Total Shares,Current Price,Gain/Loss,Gain/Loss for Hold,Max Loss for Strategy,Max Loss for Hold"
        s3 = results_file_heading(pb.StrategyNo)
        If s3.Length > 0 Then s2 &= s3
        writer1.WriteLine(s1)
        writer2.WriteLine(s2)
      End If

      total_num_shares = 0#

      Dim num_shares_bought#, current_price#, price_bought_hold#, kk%, j_dt%, j_cl%
      Dim num_shares_sold#, price_bought#, num_days_in_market%
      num_days_in_market = 0
      j_dt = 1 + pb1.date1.Count - pb.num_used
      j_cl = 1 + pb1.close.Count - pb.num_used
      price_bought_hold = pb1.close(j_cl) ' not Close(0) since the gain for the strategy starts a j = 1 not 0
      Dim num_shares_hold# = pb.initial_cash / price_bought_hold

      Dim bBought = False

      kk = 0
      Dim date_sold As Date
      date_sold = pb1.date1(j_dt)
      Dim current_cash = pb.initial_cash
      Dim cash_after_sale = pb.initial_cash
      pb.sell_date_index = -1000

      For j% = 1 To pb.num_used - 1 ' some strategies need j - 1
        j_dt = j + pb1.date1.Count - pb.num_used
        j_cl = j + pb1.close.Count - pb.num_used
        bDaySold = False
        bDayBought = False

        current_price = pb1.close(j_cl)
        bBought = True
        If total_num_shares <= 0.000001 Then bBought = False
        Dim action1%
        action1 = find_action(j, pb.StrategyNo, x, bBought, price_bought, price_sold)
        Dim days& = DateDiff(DateInterval.Day, date_sold, pb1.date1(j_dt))
        price_if_sold = pb1.close(j_cl) ' price from close the same day to be compatible with the transactions file
        amount_hold# = num_shares_hold * price_if_sold
        gain_for_hold = amount_hold - pb.initial_cash

        Select Case action1
          Case NO_BUY
            If days > 0 And pb.bIncludeInterest Then
              current_cash = cash_after_sale * (1.0 + pb.interest_rate * CDbl(days) / (365.0 * 100.0))
            End If
          Case BUY
            'price_bought = pb1.open(j_cl + 1) ' price from open the next day
            price_bought = pb1.close(j_cl) ' price from close the same day
            If days > 0 And pb.bIncludeInterest Then
              ' use 365 because days are calculated for actual days not market days
              gain_from_interest += cash_after_sale * pb.interest_rate * CDbl(days) / (365.0 * 100.0)
              current_cash = cash_after_sale * (1.0 + pb.interest_rate * CDbl(days) / (365.0 * 100.0))
            End If
            num_shares_bought = current_cash / price_bought
            bDayBought = True
            day_bought_index = j
            total_num_shares = num_shares_bought
            cost += num_shares_bought * price_bought
            current_cash = 0.0
            amount# = current_cash + total_num_shares * price_bought ' include the value of the shares
            gain1 = amount - pb.initial_cash

            If pb.bSaveResults AndAlso pb.bSaveDetails Then
              price_bought = pb1.close(j_cl) ' price from close the same day
              s1 = "buy," & pb1.date1(j_dt).ToShortDateString & "," & total_num_shares.ToString("0.000") & "," &
                price_bought.ToString("0.000") & "," & gain1.ToString("0.000") & "," & gain_for_hold.ToString("0.000")
              writer1.WriteLine(s1)
            End If

          Case STOPPED_OUT
            num_shares_sold = total_num_shares
            fraction_sold = num_shares_sold / total_num_shares
            total_num_shares = 0#
            bDaySold = True
            date_sold = pb1.date1(j_dt)
            Call sell_calculations(price_sold, num_shares_sold, fraction_sold, price_bought_hold, total_num_shares,
              cost, count_gain, count_loss, gain_for_count, loss_for_count, perc_gain_for_count, perc_loss_for_count, current_cash)
            cash_after_sale = current_cash
            amount# = current_cash + total_num_shares * price_if_sold ' include the value of the shares
            gain1 = amount - pb.initial_cash

            If pb.bSaveResults AndAlso pb.bSaveDetails Then
              s1 = "sell," & pb1.date1(j_dt).ToShortDateString & "," & total_num_shares.ToString("0.000") & "," &
                    price_sold.ToString("0.000") & "," & gain1.ToString("0.000") & "," & gain_for_hold.ToString("0.000")
              writer1.WriteLine(s1)
            End If

          Case SELL
            price_sold = pb1.close(j_cl) ' price from close the same day
            num_shares_sold = total_num_shares
            fraction_sold = num_shares_sold / total_num_shares
            total_num_shares = 0#
            bDaySold = True
            date_sold = pb1.date1(j_dt)
            Call sell_calculations(price_sold, num_shares_sold, fraction_sold, price_bought_hold, total_num_shares,
              cost, count_gain, count_loss, gain_for_count, loss_for_count, perc_gain_for_count, perc_loss_for_count, current_cash)
            cash_after_sale = current_cash
            amount# = current_cash + total_num_shares * price_if_sold ' include the value of the shares
            gain1 = amount - pb.initial_cash

            If pb.bSaveResults AndAlso pb.bSaveDetails Then
              s1 = "sell," & pb1.date1(j_dt).ToShortDateString & "," & total_num_shares.ToString("0.000") & "," &
                  price_sold.ToString("0.000") & "," & gain1.ToString("0.000") & "," & gain_for_hold.ToString("0.000")
              writer1.WriteLine(s1)
            End If
        End Select

        amount# = current_cash + total_num_shares * price_if_sold ' include the value of the shares
        gain1 = amount - pb.initial_cash
        If amount_hold - pb.initial_cash < max_loss_for_hold Then max_loss_for_hold = amount_hold - pb.initial_cash
        If amount - pb.initial_cash < max_loss Then max_loss = amount - pb.initial_cash

        If j > 1 Then
          ' Only include returns from days where the stock is owned
          If action1 = HOLD Or action1 = STOPPED_OUT Or action1 = SELL Then
            days_return = (pb1.close(j_cl) - pb1.close(j_cl - 1)) / pb1.close(j_cl - 1)
            returns.Add(days_return)
            adjusted_returns.Add(days_return - pb.interest_rate / (252.0 * 100.0))
          End If

          'If action1 = NO_BUY Or action1 = BUY Then
          '  days_return = 0.0
          '  If pb.bIncludeInterest Then days_return = pb.interest_rate / (252.0 * 100.0) ' assume return is from interest rate
          'ElseIf action1 = HOLD Or action1 = STOPPED_OUT Or action1 = SELL Then
          '  days_return = (pb1.close(j_cl) - pb1.close(j_cl - 1)) / pb1.close(j_cl - 1)
          'End If
          'adjusted_returns.Add(days_return - pb.interest_rate / (252.0 * 100.0))

          days_return_hold = (pb1.close(j_cl) - pb1.close(j_cl - 1)) / pb1.close(j_cl - 1)
          returns_hold.Add(days_return_hold)
          adjusted_returns_hold.Add(days_return_hold - pb.interest_rate / (252.0 * 100.0))
          total_value_hold.Add(amount_hold)
          total_value.Add(amount)
        End If

        If pb.bSaveResults AndAlso pb.bSaveDetails Then
          s2 = pb1.date1(j_dt).ToShortDateString & "," & total_num_shares.ToString("0.000") & "," &
            price_if_sold.ToString("0.000") & "," & gain1.ToString("0.000") & "," & gain_for_hold.ToString("0.000") & "," &
            max_loss.ToString("0.000") & "," & max_loss_for_hold.ToString("0.000")
          s3 = results_file_text(j, pb.StrategyNo)
          If s3.Length > 0 Then s2 &= s3
          writer2.WriteLine(s2)
        End If
        If (total_num_shares > 0.000001 And Not bDayBought) Or bDaySold Then num_days_in_market += 1
      Next

      If (total_num_shares <= 0.000001) Then
        Dim days& = DateDiff(DateInterval.Day, date_sold, pb1.date1.Last)
        If days > 0 And pb.bIncludeInterest Then
          gain_from_interest += current_cash * pb.interest_rate * CDbl(days) / (365.0 * 100.0)
        End If
      End If

      If pb.bSaveResults AndAlso pb.bSaveDetails Then
        writer1.Close()
        writer2.Close()
      End If

      max_drawdown_for_hold = FindMaxDrawdown(total_value_hold)
      max_drawdown = FindMaxDrawdown(total_value)

      Dim gl_ratio# = 20.0
      If count_loss > 0 And loss_for_count > 0.00001 Then gl_ratio# = gain_for_count / loss_for_count
      If count_gain > 0 Then gain_for_count = gain_for_count / CDbl(count_gain)
      If count_gain > 0 Then perc_gain_for_count = 100.0 * perc_gain_for_count / CDbl(count_gain)
      If count_loss > 0 Then loss_for_count = loss_for_count / CDbl(count_loss)
      If count_loss > 0 Then perc_loss_for_count = 100.0 * perc_loss_for_count / CDbl(count_loss)
      win_rate# = 0.0
      If count_gain + count_loss > 0 Then win_rate# = 100.0 * CDbl(count_gain) / CDbl(count_gain + count_loss)
      pb.num_trades = count_gain + count_loss
      perc_return = 100.0 * gain1 / pb.initial_cash
      perc_return_per_year = perc_return * 252.0 / CDbl(pb.num_used - 1)
      perc_return_hold = 100.0 * gain_for_hold / pb.initial_cash
      gain_per_day = 0.0
      If num_days_in_market > 0 Then gain_per_day = gain1 / CDbl(num_days_in_market)
      'perc_gain_per_day = 0.0
      'Dim perc_gain = 100.0 * gain1 / pb.initial_cash
      'perc_gain_per_day = 0.0
      'If num_days_in_market > 0 Then perc_gain_per_day = perc_gain / CDbl(num_days_in_market)
      'perc_return_per_day_hold = perc_return_hold / CDbl(pb.num_used - 1)
      gain_per_day_hold = gain_for_hold / CDbl(pb.num_used - 1)
      perc_days_in_market = 100.0 * CDbl(num_days_in_market) / CDbl(pb.num_used - 1)
      perc_of_return_for_hold = -1.0
      If gain_for_hold > 0.00001 Then perc_of_return_for_hold = 100.0 * gain1 / gain_for_hold

      ' calculate Sharpe ratio using arithmetic mean because that is how Sharpe did it
      'Dim geo_mean# = 100.0 * ((((perc_return / 100.0) + 1.0) ^ (1.0 / CDbl(pb.num_used - 2))) - 1.0)
      'Dim annual_return# = 100.0 * ((1.0 + (geo_mean / 100.0)) ^ 252 - 1.0)
      Dim std_dev# = 0.0, sharpe_ratio# = 0.0
      perc_gain_per_day = 0.0
      If returns.Count > 1 Then
        perc_gain_per_day = 100.0 * returns.Average()
        Dim adjusted_daily_return# = adjusted_returns.Average()
        ' standard deviation of adjusted_returns
        std_dev# = FindStdDev(adjusted_returns, False)
        ' Note that for a fixed interest rate, std(returns) is the same as std(returns - daily interest)
        sharpe_ratio# = System.Math.Sqrt(252.0) * adjusted_daily_return / std_dev
      End If

      perc_return_per_day_hold = 100.0 * returns_hold.Average()
      Dim adjusted_daily_return_hold# = adjusted_returns_hold.Average()
      Dim std_dev_hold# = FindStdDev(adjusted_returns_hold, False)
      Dim sharpe_ratio_hold# = System.Math.Sqrt(252.0) * adjusted_daily_return_hold / std_dev_hold

      'fi is the value to be maximized by the global optimization
      fi = pb.perc_gain_per_day_weight * perc_gain_per_day / 100.0

      ' reduce fi if the % days in market < 5
      If perc_days_in_market <= pb.min_perc_days_in_market Then
        fi -= pb.perc_days_in_market_weight * (1.0 - perc_days_in_market / pb.min_perc_days_in_market)
      End If

      ' increase fi if the % of trades with a gain increases
      fi += pb.win_rate_weight * win_rate / 100.0
      ' increase fi if the % gain/loss ratio increases
      fi += pb.gl_ratio_weight * gl_ratio
      ' increase fi if the % of gain for hold increases
      fi += pb.perc_return_per_year_weight * perc_return_per_year
      ' increase fi if the sharpe ratio increases
      fi += pb.sharpe_ratio_weight * sharpe_ratio

      ' reduce fi if the maximum loss % increases
      If max_drawdown < 0.0 Then
        fi -= pb.perc_max_drawdown_weight * max_drawdown
      End If

      Dim n%
      n = CInt(CDbl(pb.num_used + 4) * pb.min_num_trades_per_year / 252.0)
      ' reduce fi if the number of trades is less than the goal
      If n > 0 AndAlso pb.num_trades < n Then
        fi -= pb.weight_for_nt1 * (1.0 - CDbl(pb.num_trades) / CDbl(n))
      End If

      pb.num_trades_goal = n
      pb.fsave(0) = gain_per_day
      pb.fsave(1) = gain1
      pb.fsave(2) = perc_days_in_market
      pb.fsave(3) = win_rate
      pb.fsave(4) = perc_return_per_year

      If pb.bSaveResults Then
        sFileName1 = pb.summary_file
        Dim perc_gain_per_day_diff# = perc_gain_per_day - perc_return_per_day_hold
        s2 = pb.ticker & "," & pb.category & "," & pb.initial_cash & "," & (pb.num_used - 1).ToString.Trim & ","
        For i = 0 To pb.num_variables - 1
          Dim xx#
          xx = x(i)
          If pb.x_type(i) = "i" Then xx = System.Math.Round(x(i))
          s2 = s2 & xx.ToString("0.000") & ","
        Next

        Dim s4 = ""
        If perc_of_return_for_hold > 0.0 Then s4 = perc_of_return_for_hold.ToString("0.00")
        s2 = s2 &
          perc_days_in_market.ToString("0.00") & "," & (count_gain + count_loss).ToString("0") & "," & win_rate.ToString("0.00") & "," &
          gain_per_day.ToString("0.0000") & "," & gain_per_day_hold.ToString("0.0000") & "," & perc_gain_for_count.ToString("0.000") & "," & perc_loss_for_count.ToString("0.000") & "," & gain1.ToString("0.000") & "," & gain_for_hold.ToString("0.000") & "," &
          max_drawdown.ToString("0.00") & "," & max_drawdown_for_hold.ToString("0.00") & "," & perc_gain_per_day.ToString("0.0000") & "," & perc_gain_per_day_diff.ToString("0.00000") & "," & gl_ratio.ToString("0.000") & "," &
          s4 & "," & perc_return.ToString("0.00") & "," & perc_return_hold.ToString("0.00") & "," & sharpe_ratio.ToString("0.000") & "," & sharpe_ratio_hold.ToString("0.000")

        If File.Exists(sFileName1) Then
          Dim writer3 As New StreamWriter(sFileName1, True)
          writer3.WriteLine(s2)
          writer3.Close()
        Else
          Dim writer3 As New StreamWriter(sFileName1)
          s1 = "Ticker,Category,Initial Cash,Num of Days,"
          For i = 0 To pb.num_variables - 1
            s1 = s1 & "x(" & i.ToString.Trim & "),"
          Next
          s1 &=
            "Days in Market %,Num of trades,Win rate %,Gain per day,Gain per day hold,Average gain %,Average loss %,Final gain,Gain for hold,Max drawdown %,Max drawdown % for hold,% Gain per day,% Gain per day diff,Gain/loss ratio,% of return for hold,Return % ,Return % for hold,Sharpe ratio,Sharpe ratio for hold"
          writer3.WriteLine(s1)
          writer3.WriteLine(s2)
          writer3.Close()
        End If
      End If
    Catch e As Exception
      MessageBox.Show("Error writing output files: " & e.Message)
      RunStrategy1 = -2
      Exit Function
    End Try

    If pb.bSaveResults And pb.bDisplayMessage Then
      s2 = "Win rate % = " & win_rate.ToString("0.00") & "  Average % gain = " & perc_gain_for_count.ToString("0.00") & "  Average % loss = " & perc_loss_for_count.ToString("0.00") & Environment.NewLine &
        "Gain per day = " & gain_per_day.ToString("0.0000") & "  Gain per day hold = " & gain_per_day_hold.ToString("0.0000") & Environment.NewLine &
        "Final gain = " & gain1.ToString("0.000") & "  Gain for hold = " & gain_for_hold.ToString("0.000") & Environment.NewLine &
        "Max drawdown % = " & max_drawdown.ToString("0.00") & "  Max drawdown % hold = " & max_drawdown_for_hold.ToString("0.00") & Environment.NewLine &
        "Number of trades = " & (count_gain + count_loss).ToString("0") & "    Days in market % = " & perc_days_in_market.ToString("0.00") & Environment.NewLine &
        "Gain from interest = " & gain_from_interest.ToString("0.00")
      MessageBox.Show(s2)
    End If
    RunStrategy1 = 0
  End Function
  Sub sell_calculations(price_sold#, num_shares_sold#, fraction_sold#, price_bought_hold#, total_num_shares#,
    ByRef cost#, ByRef count_gain%,
     ByRef count_loss%, ByRef gain_for_count#, ByRef loss_for_count#, ByRef perc_gain_for_count#, ByRef perc_loss_for_count#, ByRef current_cash#)
    Dim cost1#, gain_for_this_sale#
    cost1 = cost * fraction_sold
    gain_for_this_sale = num_shares_sold * price_sold - cost1
    If gain_for_this_sale > 0.0 Then
      count_gain += 1
      gain_for_count += gain_for_this_sale
      perc_gain_for_count += gain_for_this_sale / cost1
    Else
      count_loss += 1
      loss_for_count += System.Math.Abs(gain_for_this_sale)
      perc_loss_for_count += System.Math.Abs(gain_for_this_sale / cost1)
    End If
    current_cash += num_shares_sold * price_sold
    cost -= cost1
    If total_num_shares <= 0 Then cost = 0.0
  End Sub
  Function FindStdDev#(x As List(Of Double), as_sample As Boolean)
    FindStdDev = 0.0
    Dim mean# = x.Average()

    Dim squares_query =
        From value In x
        Select (value - mean) * (value - mean)

    Dim sum_squares# = squares_query.Sum()

    If (as_sample) Then
      Return System.Math.Sqrt(sum_squares / (x.Count() - 1))
    Else
      Return System.Math.Sqrt(sum_squares / x.Count())
    End If
  End Function
  Function FindMaxDrawdown#(amounts As List(Of Double))
    FindMaxDrawdown = 0.0
    Dim peak# = -100.0
    Dim maxDrawdown# = 0.0

    For Each amount In amounts
      If (amount >= peak) Then
        peak = amount
      Else
        Dim drawdown# = 1.0 - amount / peak
        maxDrawdown = System.Math.Max(maxDrawdown, drawdown)
      End If
    Next
    FindMaxDrawdown = maxDrawdown * 100.0
  End Function


  Function FindRmi%(quotes As IEnumerable(Of Quote), n%, m%, ns%, ByRef date1 As List(Of Date), ByRef rmi As List(Of Double), ByRef signal As List(Of Double))
    ' n%...the number of periods used for the smoothing
    ' m%...the interval (expressed as number of periods) used to find the change in price
    ' ns%...the number of periods used for the EMA signal line calculation

    FindRmi = -1
    Dim i%, L%
    Dim up#(), down#(), rmi1#(), diff#

    Dim date11 As New List(Of Date)
    Dim close As New List(Of Double)
    Call GetQuoteCloseLists(quotes, date11, close)
    Dim x As Array = close.ToArray
    L = x.Length
    If L < 2 * n + ns + m Then
      MessageBox.Show("FindRmi --- Not enough points")
      Exit Function
    End If

    ReDim up#(0 To L - m - 1), down#(0 To L - m - 1)
    For i = m To L - 1
      up(i - m) = 0#
      down(i - m) = 0#
      diff = x(i) - x(i - m)
      If diff > 0 Then
        up(i - m) = diff
      ElseIf (diff < 0) Then
        down(i - m) = System.Math.Abs(diff)
      End If
    Next
    ' initialize with the SMA
    Dim sma_up#, sma_down#, smooth_up#, smooth_down#, dN#, multiplier#, dNs#, sma#, ema#
    sma_up = 0#
    sma_down = 0#

    dN = CDbl(n)
    For i = 0 To n - 1
      sma_up = sma_up + up(i)
      sma_down = sma_down + down(i)
    Next
    smooth_up = sma_up / dN
    smooth_down = sma_down / dN

    ' continue with the smooth


    ReDim rmi1#(0 To L - m - n - 1)
    multiplier = 1.0# / dN
    For i = n + m To L - 1
      smooth_up = up(i - m) * multiplier + (1.0# - multiplier) * smooth_up
      smooth_down = down(i - m) * multiplier + (1.0# - multiplier) * smooth_down
      If smooth_up + smooth_down <= 0.0000000001 Then
        rmi1(i - n - m) = 50.0#
      Else
        rmi1(i - n - m) = 100.0# - 100.0# * smooth_down / (smooth_down + smooth_up)
      End If
    Next

    'signal line
    dNs = CDbl(ns)
    sma = 0.0
    For i = 0 To ns - 1
      sma = sma + rmi1(i)
    Next
    ema = sma / dNs

    date1.Clear()
    rmi.Clear()
    signal.Clear()
    multiplier = 2.0# / (dNs + 1.0)
    For i = n + ns + m To L - 1
      ema = rmi1(i - n - m) * multiplier + (1.0# - multiplier) * ema
      pb1.date1.Add(date11(i))
      rmi.Add(rmi1(i - n - m))
      signal.Add(ema)
    Next
    FindRmi = 0
  End Function
  Sub GetEmaLists(result As IEnumerable(Of EmaResult), ByRef date1 As List(Of Date), ByRef ema As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Ema
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    ema = (From x In result
           Select x1 = x.Ema
           Where x1 IsNot Nothing
           Select CDbl(x1)).ToList
  End Sub
  Function GetQuotes(max_num_points%, ticker$, connection_string$) As List(Of Skender.Stock.Indicators.Quote)
    Dim query1$
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!

    Dim quotes As New List(Of Skender.Stock.Indicators.Quote)
    quotes.Clear()
    GetQuotes = quotes
    cn.ConnectionString = connection_string
    ' I want the BOTTOM records of the original table in ascending order
    query1 = "Select * FROM (Select TOP " & Trim$(Str$(max_num_points)) & " * FROM market_price t1 WHERE Ticker='" & ticker & "' ORDER BY t1.Date DESC) t2 ORDER BY t2.Date ASC"
    Try
      Dim sda As New SqlDataAdapter(query1, cn)
      Dim dt As DataTable = New DataTable
      sda.Fill(dt)
      If dt.Rows.Count > 0 Then
        quotes = (From x In dt.AsEnumerable()
                  Select date1 = x.Field(Of Int32)("Date"), high1 = x.Field(Of Decimal)("High"), low1 = x.Field(Of Decimal)("Low"),
                    open1 = x.Field(Of Decimal)("Open"), close1 = x.Field(Of Decimal)("Close"), volume1 = x.Field(Of Long)("Volume")
                  Select New Skender.Stock.Indicators.Quote With {
                  .[Date] = ConvertDate(date1),
                  .High = CDbl(high1),
                  .Low = CDbl(low1),
                  .Open = CDbl(open1),
                  .Close = CDbl(close1),
                  .Volume = CDbl(volume1)}
                 ).ToList
      End If
    Catch e As Exception
      MessageBox.Show(e.Message)
      Exit Function
    End Try
    GetQuotes = quotes
  End Function

  Function ConvertDate$(date1&)
    Dim s1$
    s1 = date1.ToString.Trim
    Dim parsedDate = DateTime.ParseExact(s1, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
    Dim formattedDate = parsedDate.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture)
    ConvertDate = formattedDate
  End Function
  Sub GetQuoteLists(quotes As IEnumerable(Of Quote), ByRef date1 As List(Of Date), ByRef high As List(Of Double), ByRef low As List(Of Double),
                         ByRef open As List(Of Double), ByRef close As List(Of Double), ByRef volume As List(Of Double))
    date1 = (From x In quotes
             Select date_value = x.[Date]
             Select CDate(date_value)).ToList

    high = (From x In quotes
            Select x1 = x.High
            Select CDbl(x1)).ToList

    low = (From x In quotes
           Select x1 = x.Low
           Select CDbl(x1)).ToList

    open = (From x In quotes
            Select x1 = x.Open
            Select CDbl(x1)).ToList

    close = (From x In quotes
             Select x1 = x.Close
             Select CDbl(x1)).ToList

    volume = (From x In quotes
              Select x1 = x.Volume
              Select CDbl(x1)).ToList
  End Sub
  Sub GetHeikinAshiLists(result As IEnumerable(Of HeikinAshiResult), ByRef date1 As List(Of Date), ByRef high As List(Of Double), ByRef low As List(Of Double),
                         ByRef open As List(Of Double), ByRef close As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date]
             Select CDate(date_value)).ToList

    high = (From x In result
            Select x1 = x.High
            Select CDbl(x1)).ToList

    low = (From x In result
           Select x1 = x.Low
           Select CDbl(x1)).ToList

    open = (From x In result
            Select x1 = x.Open
            Select CDbl(x1)).ToList

    close = (From x In result
             Select x1 = x.Close
             Select CDbl(x1)).ToList
  End Sub

  Sub GetKeltnerLists(result As IEnumerable(Of KeltnerResult), ByRef date1 As List(Of Date), ByRef centerLine As List(Of Double), ByRef upperBand As List(Of Double),
                         ByRef lowerBand As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Centerline
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    centerLine = (From x In result
                  Select x1 = x.Centerline
                  Where x1 IsNot Nothing
                  Select CDbl(x1)).ToList

    upperBand = (From x In result
                 Select x1 = x.UpperBand
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList

    lowerBand = (From x In result
                 Select x1 = x.LowerBand
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList
  End Sub

  Sub GetQuoteCloseLists(quotes As IEnumerable(Of Quote), ByRef date1 As List(Of Date), ByRef close As List(Of Double))
    date1 = (From x In quotes
             Select date_value = x.[Date]
             Select CDate(date_value)).ToList

    close = (From x In quotes
             Select x1 = x.Close
             Select CDbl(x1)).ToList
  End Sub


  Sub GetQuoteVolumeLists(quotes As IEnumerable(Of Quote), ByRef date1 As List(Of Date), ByRef volume As List(Of Double))
    date1 = (From x In quotes
             Select date_value = x.[Date]
             Select CDate(date_value)).ToList

    volume = (From x In quotes
              Select x1 = x.Volume
              Select CDbl(x1)).ToList
  End Sub
  Sub GetBollingerLists(result As IEnumerable(Of BollingerBandsResult), ByRef date1 As List(Of Date), ByRef sma As List(Of Double), ByRef upperBand As List(Of Double),
                         ByRef lowerBand As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Sma
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    sma = (From x In result
           Select x1 = x.Sma
           Where x1 IsNot Nothing
           Select CDbl(x1)).ToList

    upperBand = (From x In result
                 Select x1 = x.UpperBand
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList

    lowerBand = (From x In result
                 Select x1 = x.LowerBand
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList
  End Sub

  Sub GetDonchianLists(result As IEnumerable(Of DonchianResult), ByRef date1 As List(Of Date), ByRef centerLine As List(Of Double), ByRef upperBand As List(Of Double),
                         ByRef lowerBand As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Centerline
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    centerLine = (From x In result
                  Select x1 = x.Centerline
                  Where x1 IsNot Nothing
                  Select CDbl(x1)).ToList

    upperBand = (From x In result
                 Select x1 = x.UpperBand
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList

    lowerBand = (From x In result
                 Select x1 = x.LowerBand
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList
  End Sub
  Sub GetSmaLists(result As IEnumerable(Of SmaResult), ByRef date1 As List(Of Date), ByRef sma As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Sma
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    sma = (From x In result
           Select x1 = x.Sma
           Where x1 IsNot Nothing
           Select CDbl(x1)).ToList
  End Sub
  Sub GetRsiLists(result As IEnumerable(Of RsiResult), ByRef date1 As List(Of Date), ByRef rsi As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Rsi
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    rsi = (From x In result
           Select x1 = x.Rsi
           Where x1 IsNot Nothing
           Select CDbl(x1)).ToList
  End Sub
  Sub GetMfiLists(result As IEnumerable(Of MfiResult), ByRef date1 As List(Of Date), ByRef mfi As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Mfi
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    mfi = (From x In result
           Select x1 = x.Mfi
           Where x1 IsNot Nothing
           Select CDbl(x1)).ToList
  End Sub
  Sub GetMacdLists(result As IEnumerable(Of MacdResult), ByRef date1 As List(Of Date), ByRef macd As List(Of Double), ByRef signal As List(Of Double),
    ByRef histogram As List(Of Double))

    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Macd
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    macd = (From x In result
            Select x1 = x.Macd
            Where x1 IsNot Nothing
            Select CDbl(x1)).ToList

    signal = (From x In result
              Select x1 = x.Signal
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList

    histogram = (From x In result
                 Select x1 = x.Histogram
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList
  End Sub
  Sub GetObvLists(result As IEnumerable(Of ObvResult), ByRef lstDate As List(Of Date), ByRef lstObv As List(Of Double), ByRef lstObvSma As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.ObvSma
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstObv = (From x In result
              Select obv_value = x.Obv, x1 = x.ObvSma
              Where x1 IsNot Nothing
              Select CDbl(obv_value)).ToList

    lstObvSma = (From x In result
                 Select x1 = x.ObvSma
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList
  End Sub
  Sub GetParabolicSarLists(result As IEnumerable(Of ParabolicSarResult), ByRef date1 As List(Of Date), ByRef sar As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.Sar
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    sar = (From x In result
           Select x1 = x.Sar
           Where x1 IsNot Nothing
           Select CDbl(x1)).ToList
  End Sub

  Sub GetStochRsiLists(result As IEnumerable(Of StochRsiResult), ByRef date1 As List(Of Date), ByRef stochRsi As List(Of Double), ByRef signal As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.StochRsi
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    stochRsi = (From x In result
                Select x1 = x.StochRsi
                Where x1 IsNot Nothing
                Select CDbl(x1)).ToList

    signal = (From x In result
              Select x1 = x.Signal
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList
  End Sub
  Sub GetStdDevLists(result As IEnumerable(Of StdDevResult), ByRef date1 As List(Of Date), ByRef stdDev As List(Of Double), ByRef zScore As List(Of Double), ByRef mean As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.StdDev
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    stdDev = (From x In result
              Select x1 = x.StdDev
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList

    zScore = (From x In result
              Select x1 = x.ZScore
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList

    mean = (From x In result
            Select x1 = x.Mean
            Where x1 IsNot Nothing
            Select CDbl(x1)).ToList
  End Sub

  Sub GetSuperTrendLists(result As IEnumerable(Of SuperTrendResult), ByRef date1 As List(Of Date), ByRef superTrend As List(Of Double))
    date1 = (From x In result
             Select date_value = x.[Date], x1 = x.SuperTrend
             Where x1 IsNot Nothing
             Select CDate(date_value)).ToList

    superTrend = (From x In result
                  Select x1 = x.SuperTrend
                  Where x1 IsNot Nothing
                  Select CDbl(x1)).ToList
  End Sub
  Sub GetSuperTrendLists1(result As IEnumerable(Of SuperTrendResult), ByRef superTrend As List(Of Double), ByRef isUpTrend As List(Of Boolean))
    superTrend = (From x In result
                  Select x1 = x.SuperTrend
                  Where x1 IsNot Nothing
                  Select CDbl(x1)).ToList

    isUpTrend = (From x In result
                 Select x1 = x.LowerBand, x2 = x.SuperTrend
                 Where x2 IsNot Nothing
                 Select CBool(x1 IsNot Nothing)).ToList
  End Sub

  Function SumOfList(x As List(Of Double), n%) As List(Of Double)
    SumOfList = Nothing
    Dim xx As New List(Of Double)
    Dim i%, j%, sum#
    For i = n - 1 To x.Count - 1
      sum = 0.0
      For j = i - n + 1 To i
        sum += x(j)
      Next
      xx.Add(sum)
    Next
    SumOfList = xx
  End Function

  Function SmaOfList(x As List(Of Double), n%) As List(Of Double)
    SmaOfList = Nothing
    Dim xx As New List(Of Double)
    Dim i%, j%, sum#
    For i = n - 1 To x.Count - 1
      sum = 0.0
      For j = i - n + 1 To i
        sum += x(j)
      Next
      xx.Add(sum / CDbl(n))
    Next
    SmaOfList = xx
  End Function

  Function MultiplyForList(a As List(Of Double), b As List(Of Double), n%) As List(Of Double)
    MultiplyForList = Nothing
    Dim ab As New List(Of Double)
    Dim i%
    If a.Count < n Then Exit Function
    If b.Count < n Then Exit Function
    For i = 0 To n - 1
      ab.Add(a(i) * b(i))
    Next
    MultiplyForList = ab
  End Function
  Function ReadExcelFileSAX(fileName$) As DataTable
    ReadExcelFileSAX = Nothing
    Dim dt As New DataTable
    Try
      Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
        Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
        Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()

        Dim reader As OpenXmlReader = OpenXmlReader.Create(worksheetPart)
        Dim count% = 0
        While reader.Read()
          If reader.ElementType = GetType(Row) Then
            Dim dr = dt.NewRow()
            'if row has attribute then it is not an empty row
            If (reader.HasAttributes) Then
              'read the child of row element which is cells
              'here first element
              reader.ReadFirstChild()
              Do
                'find xml cell element type 
                If (reader.ElementType = GetType(Cell)) Then
                  Dim c As Cell = reader.LoadCurrentElement()
                  Dim CellValue$
                  Dim actualCellIndex% = CellReferenceToIndex(c)

                  If ((Not IsNothing(c.DataType)) AndAlso c.DataType.Equals(CellValues.SharedString)) Then
                    Dim ssi As SharedStringItem = workbookPart.SharedStringTablePart.SharedStringTable.Elements().ElementAt(Int32.Parse(c.CellValue.InnerText))
                    CellValue = ssi.Text.Text.Trim
                  Else
                    CellValue = c.CellValue.InnerText.Trim
                    'if row index Is 0 it is header so columns headers are added and also can do some headers check incase
                  End If

                  If (count = 0) Then
                    dt.Columns.Add(CellValue)
                  Else
                    ' instead of dr(c.CellReference) = cellValue
                    dr(actualCellIndex) = CellValue
                  End If
                End If
              Loop While (reader.ReadNextSibling())

              'if it is not the header row then append rowdata to the datatable
              If (count <> 0) Then
                dt.Rows.Add(dr)
              End If
              count += 1
            End If
          End If
        End While
      End Using
    Catch e As Exception
      MessageBox.Show("Error reading Excel file " & fileName & ": " & e.Message)
      Exit Function
    End Try
    ReadExcelFileSAX = dt
  End Function

  Function ReadExcelFileSAX1(fileName$) As Array
    ' This version can be used when different rows have different numbers of columns
    ReadExcelFileSAX1 = Nothing
    Dim lstCellInfo As New List(Of CellInfo)
    Dim values(,)
    Dim MaxRow% = 0, MaxCol% = 0
    Try
      Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
        Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
        Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts.First()

        Dim reader As OpenXmlReader = OpenXmlReader.Create(worksheetPart)
        Dim count% = 0
        While reader.Read()
          If reader.ElementType = GetType(Row) Then
            'if row has attribute then it is not an empty row
            If (reader.HasAttributes) Then
              'read the child of row element which is cells
              'here first element
              reader.ReadFirstChild()
              Do
                'find xml cell element type 
                If (reader.ElementType = GetType(Cell)) Then
                  Dim c As Cell = reader.LoadCurrentElement()
                  Dim CellValue$
                  Dim actualCellIndex% = CellReferenceToIndex(c)

                  If ((Not IsNothing(c.DataType)) AndAlso c.DataType.Equals(CellValues.SharedString)) Then
                    Dim ssi As SharedStringItem = workbookPart.SharedStringTablePart.SharedStringTable.Elements().ElementAt(Int32.Parse(c.CellValue.InnerText))
                    CellValue = ssi.Text.Text.Trim
                  Else
                    CellValue = c.CellValue.InnerText.Trim
                    'if row index Is 0 it is header so columns headers are added and also can do some headers check incase
                  End If

                  Dim ci As New CellInfo(count, actualCellIndex, CellValue)
                  lstCellInfo.Add(ci)
                  If count > MaxRow Then MaxRow = count
                  If actualCellIndex > MaxCol Then MaxCol = actualCellIndex
                End If
              Loop While (reader.ReadNextSibling())
              count += 1
            End If
          End If
        End While
      End Using
    Catch e As Exception
      MessageBox.Show("Error reading Excel file " & fileName & ": " & e.Message)
      Exit Function
    End Try

    Dim i%, j%
    ReDim values(0 To MaxRow, 0 To MaxCol)
    For i = 0 To MaxRow
      For j = 0 To MaxCol
        values(i, j) = ""
      Next
    Next

    For i = 0 To lstCellInfo.Count - 1
      Dim ci = lstCellInfo(i)
      values(ci.row, ci.column) = ci.value
    Next
    ReadExcelFileSAX1 = values
  End Function

  Function CellReferenceToIndex(cell As Cell)
    Dim i%
    Dim index% = 0
    Dim reference$ = cell.CellReference.ToString().ToUpper().Trim
    If Len(reference) <= 0 Then Return index
    For i = 1 To Len(reference)
      Dim ch As Char = Mid(reference, i, 1)
      If Char.IsLetter(ch) Then
        Dim value1% = Asc(ch) - Asc("A")
        If i = 1 Then
          index = value1
        Else
          index = ((index + 1) * 26) + value1
        End If
      Else
        Return index
      End If
    Next
    Return index
  End Function
  Function ReadConnectionString(ByVal sFileName$)
    ReadConnectionString = 0
    If (Dir(sFileName$) = "") Then Exit Function
    If Not File.Exists(sFileName) Then Exit Function
    Dim line$
    ReadConnectionString = -1
    line = ""

    Try
      Dim reader As New StreamReader(sFileName)
      With UserInput
        While (Not reader.EndOfStream)
          line = reader.ReadLine()
          If (line Is Nothing) Then
            reader.Close()
            Exit Function
          End If
          line = line.Trim
          If line.Length <= 0 Then
            reader.Close()
            Exit Function
          End If
          Dim items = line.Split(",")
          Select Case (Trim$(items(0)))
            Case "connection_string"
              .connection_string = items(1).Trim
          End Select
        End While
      End With
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & sFileName & ": " & e.Message)
      ReadConnectionString = -2
      Exit Function
    End Try
    ReadConnectionString = 0
  End Function
  Sub InitializeDefaults()
    With UserInput
      'example: connection_string = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
      .connection_string = "The connection string goes here"
      .ticker = ""
      .category = ""
      .num_for_calc = 0
      .strategy_no = 1
      .num_of_attempts = 10
      .max_solver_iterations = 90000
      .folder_name = ""
      .initial_cash = 1000.0
      .interest_rate = 3.0
      .include_interest = "True"
    End With
  End Sub

  Function ReadDefaults(ByVal sFileName$)
    ReadDefaults = 0
    If (Dir(sFileName$) = "") Then Exit Function
    If Not File.Exists(sFileName) Then Exit Function
    Dim line$
    ReadDefaults = -1
    line = ""

    Try
      Dim reader As New StreamReader(sFileName)
      With UserInput
        While (Not reader.EndOfStream)
          line = reader.ReadLine()
          If (line Is Nothing) Then
            reader.Close()
            Exit Function
          End If
          line = line.Trim
          If line.Length <= 0 Then
            reader.Close()
            Exit Function
          End If
          Dim items = line.Split(",")
          Select Case (Trim$(items(0)))
            Case "ticker"
              .ticker = items(1).Trim
            Case "category"
              .category = items(1).Trim
            Case "num_for_calc"
              .num_for_calc = CInt(items(1).Trim)
            Case "strategy_no"
              .strategy_no = CInt(items(1).Trim)
            Case "num_of_attempts"
              .num_of_attempts = CInt(items(1).Trim)
            Case "max_solver_iterations"
              .max_solver_iterations = CInt(items(1).Trim)
            Case "folder_name"
              .folder_name = items(1).Trim
            Case "initial_cash"
              .initial_cash = CDbl(items(1).Trim)
            Case "interest_rate"
              .interest_rate = CDbl(items(1).Trim)
            Case "include_interest"
              .include_interest = items(1).Trim
          End Select
        End While
      End With
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & sFileName & ": " & e.Message)
      ReadDefaults = -2
      Exit Function
    End Try
    ReadDefaults = 0
  End Function
  Function SaveDefaults(ByVal sFileName$)
    SaveDefaults = -1
    If File.Exists(sFileName) Then File.Delete(sFileName)
    Try
      Dim writer1 As New StreamWriter(sFileName)
      With UserInput
        writer1.WriteLine("ticker," & .ticker.Trim)
        writer1.WriteLine("category," & .category.Trim)
        writer1.WriteLine("num_for_calc," & .num_for_calc.ToString.Trim)
        writer1.WriteLine("strategy_no," & .strategy_no.ToString.Trim)
        writer1.WriteLine("num_of_attempts," & .num_of_attempts.ToString.Trim)
        writer1.WriteLine("max_solver_iterations," & .max_solver_iterations.ToString.Trim)
        writer1.WriteLine("folder_name," & .folder_name.Trim)
        writer1.WriteLine("initial_cash," & .initial_cash.ToString.Trim)
        writer1.WriteLine("interest_rate," & .interest_rate.ToString.Trim)
        writer1.WriteLine("include_interest," & .include_interest.Trim)
      End With
      writer1.Close()
    Catch e As Exception
      MessageBox.Show("Error writing file " & sFileName & ": " & e.Message)
      SaveDefaults = -2
      Exit Function
    End Try
    SaveDefaults = 0
  End Function
  Function ResizeLists(num_for_calc%, ByRef date1 As List(Of Date), ByRef list0 As List(Of Double), ByRef Optional list1 As List(Of Double) = Nothing,
    ByRef Optional list2 As List(Of Double) = Nothing, ByRef Optional list3 As List(Of Double) = Nothing, ByRef Optional list4 As List(Of Double) = Nothing,
    ByRef Optional list5 As List(Of Double) = Nothing, ByRef Optional list6 As List(Of Double) = Nothing, ByRef Optional list7 As List(Of Double) = Nothing,
    ByRef Optional list8 As List(Of Double) = Nothing, ByRef Optional list9 As List(Of Double) = Nothing, ByRef Optional list10 As List(Of Double) = Nothing,
    ByRef Optional list11 As List(Of Double) = Nothing, ByRef Optional list12 As List(Of Double) = Nothing, ByRef Optional list13 As List(Of Double) = Nothing,
    ByRef Optional list14 As List(Of Double) = Nothing, ByRef Optional list15 As List(Of Double) = Nothing, ByRef Optional list16 As List(Of Double) = Nothing)
    Dim min_num_points%
    ResizeLists = -1

    min_num_points = date1.Count
    If list0.Count < min_num_points Then min_num_points = list0.Count
    If Not IsNothing(list1) Then
      If list1.Count < min_num_points Then min_num_points = list1.Count
    End If
    If Not IsNothing(list2) Then
      If list2.Count < min_num_points Then min_num_points = list2.Count
    End If
    If Not IsNothing(list3) Then
      If list3.Count < min_num_points Then min_num_points = list3.Count
    End If
    If Not IsNothing(list4) Then
      If list4.Count < min_num_points Then min_num_points = list4.Count
    End If
    If Not IsNothing(list5) Then
      If list5.Count < min_num_points Then min_num_points = list5.Count
    End If
    If Not IsNothing(list6) Then
      If list6.Count < min_num_points Then min_num_points = list6.Count
    End If
    If Not IsNothing(list7) Then
      If list7.Count < min_num_points Then min_num_points = list7.Count
    End If
    If Not IsNothing(list8) Then
      If list8.Count < min_num_points Then min_num_points = list8.Count
    End If
    If Not IsNothing(list9) Then
      If list9.Count < min_num_points Then min_num_points = list9.Count
    End If
    If Not IsNothing(list10) Then
      If list10.Count < min_num_points Then min_num_points = list10.Count
    End If
    If Not IsNothing(list11) Then
      If list11.Count < min_num_points Then min_num_points = list11.Count
    End If
    If Not IsNothing(list12) Then
      If list12.Count < min_num_points Then min_num_points = list12.Count
    End If
    If Not IsNothing(list13) Then
      If list13.Count < min_num_points Then min_num_points = list13.Count
    End If
    If Not IsNothing(list14) Then
      If list14.Count < min_num_points Then min_num_points = list14.Count
    End If
    If Not IsNothing(list15) Then
      If list15.Count < min_num_points Then min_num_points = list15.Count
    End If
    If Not IsNothing(list16) Then
      If list16.Count < min_num_points Then min_num_points = list16.Count
    End If

    If min_num_points < 10 Then Exit Function

    ResizeListOfDate(min_num_points, num_for_calc, date1)
    ResizeListOfDbl(min_num_points, num_for_calc, list0)
    If Not IsNothing(list1) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list1)
    End If
    If Not IsNothing(list2) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list2)
    End If
    If Not IsNothing(list3) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list3)
    End If
    If Not IsNothing(list4) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list4)
    End If
    If Not IsNothing(list5) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list5)
    End If
    If Not IsNothing(list6) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list6)
    End If
    If Not IsNothing(list7) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list7)
    End If
    If Not IsNothing(list8) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list8)
    End If
    If Not IsNothing(list9) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list9)
    End If
    If Not IsNothing(list10) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list10)
    End If
    If Not IsNothing(list11) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list11)
    End If
    If Not IsNothing(list12) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list12)
    End If
    If Not IsNothing(list13) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list13)
    End If
    If Not IsNothing(list14) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list14)
    End If
    If Not IsNothing(list15) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list15)
    End If
    If Not IsNothing(list16) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list16)
    End If

    ResizeLists = 0
  End Function
  Sub ResizeLists1(num_for_calc%, ByRef list0 As List(Of Double), ByRef Optional list1 As List(Of Double) = Nothing,
    ByRef Optional list2 As List(Of Double) = Nothing, ByRef Optional list3 As List(Of Double) = Nothing, ByRef Optional list4 As List(Of Double) = Nothing,
    ByRef Optional list5 As List(Of Double) = Nothing, ByRef Optional list6 As List(Of Double) = Nothing, ByRef Optional list7 As List(Of Double) = Nothing,
    ByRef Optional list8 As List(Of Double) = Nothing, ByRef Optional list9 As List(Of Double) = Nothing, ByRef Optional list10 As List(Of Double) = Nothing,
    ByRef Optional list11 As List(Of Double) = Nothing, ByRef Optional list12 As List(Of Double) = Nothing, ByRef Optional list13 As List(Of Double) = Nothing,
    ByRef Optional list14 As List(Of Double) = Nothing, ByRef Optional list15 As List(Of Double) = Nothing, ByRef Optional list16 As List(Of Double) = Nothing)
    Dim min_num_points%
    min_num_points = list0.Count
    If Not IsNothing(list1) Then
      If list1.Count < min_num_points Then min_num_points = list1.Count
    End If
    If Not IsNothing(list2) Then
      If list2.Count < min_num_points Then min_num_points = list2.Count
    End If
    If Not IsNothing(list3) Then
      If list3.Count < min_num_points Then min_num_points = list3.Count
    End If
    If Not IsNothing(list4) Then
      If list4.Count < min_num_points Then min_num_points = list4.Count
    End If
    If Not IsNothing(list5) Then
      If list5.Count < min_num_points Then min_num_points = list5.Count
    End If
    If Not IsNothing(list6) Then
      If list6.Count < min_num_points Then min_num_points = list6.Count
    End If
    If Not IsNothing(list7) Then
      If list7.Count < min_num_points Then min_num_points = list7.Count
    End If
    If Not IsNothing(list8) Then
      If list8.Count < min_num_points Then min_num_points = list8.Count
    End If
    If Not IsNothing(list9) Then
      If list9.Count < min_num_points Then min_num_points = list9.Count
    End If
    If Not IsNothing(list10) Then
      If list10.Count < min_num_points Then min_num_points = list10.Count
    End If
    If Not IsNothing(list11) Then
      If list11.Count < min_num_points Then min_num_points = list11.Count
    End If
    If Not IsNothing(list12) Then
      If list12.Count < min_num_points Then min_num_points = list12.Count
    End If
    If Not IsNothing(list13) Then
      If list13.Count < min_num_points Then min_num_points = list13.Count
    End If
    If Not IsNothing(list14) Then
      If list14.Count < min_num_points Then min_num_points = list14.Count
    End If
    If Not IsNothing(list15) Then
      If list15.Count < min_num_points Then min_num_points = list15.Count
    End If
    If Not IsNothing(list16) Then
      If list16.Count < min_num_points Then min_num_points = list16.Count
    End If

    ResizeListOfDbl(min_num_points, num_for_calc, list0)
    If Not IsNothing(list1) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list1)
    End If
    If Not IsNothing(list2) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list2)
    End If
    If Not IsNothing(list3) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list3)
    End If
    If Not IsNothing(list4) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list4)
    End If
    If Not IsNothing(list5) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list5)
    End If
    If Not IsNothing(list6) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list6)
    End If
    If Not IsNothing(list7) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list7)
    End If
    If Not IsNothing(list8) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list8)
    End If
    If Not IsNothing(list9) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list9)
    End If
    If Not IsNothing(list10) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list10)
    End If
    If Not IsNothing(list11) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list11)
    End If
    If Not IsNothing(list12) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list12)
    End If
    If Not IsNothing(list13) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list13)
    End If
    If Not IsNothing(list14) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list14)
    End If
    If Not IsNothing(list15) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list15)
    End If
    If Not IsNothing(list16) Then
      ResizeListOfDbl(min_num_points, num_for_calc, list16)
    End If
  End Sub
  Sub ResizeListOfDate(min_num_points%, num_for_calc%, ByRef lstList As List(Of Date))
    Dim num_elements%
    num_elements = lstList.Count
    If num_elements > min_num_points Then
      lstList.RemoveRange(0, num_elements - min_num_points)
      num_elements = min_num_points
    End If
    If num_elements > num_for_calc Then
      lstList.RemoveRange(0, num_elements - num_for_calc)
    End If
  End Sub
  Sub ResizeListOfDbl(min_num_points%, num_for_calc%, ByRef lstList As List(Of Double))
    Dim num_elements%
    num_elements = lstList.Count
    If num_elements > min_num_points Then
      lstList.RemoveRange(0, num_elements - min_num_points)
      num_elements = min_num_points
    End If
    If num_elements > num_for_calc Then
      lstList.RemoveRange(0, num_elements - num_for_calc)
    End If
  End Sub

  Function MinNumPoints(num_for_calc%, ByRef date1 As List(Of Date), ByRef list0 As List(Of Double), ByRef Optional list1 As List(Of Double) = Nothing,
    ByRef Optional list2 As List(Of Double) = Nothing, ByRef Optional list3 As List(Of Double) = Nothing, ByRef Optional list4 As List(Of Double) = Nothing,
    ByRef Optional list5 As List(Of Double) = Nothing, ByRef Optional list6 As List(Of Double) = Nothing, ByRef Optional list7 As List(Of Double) = Nothing,
    ByRef Optional list8 As List(Of Double) = Nothing, ByRef Optional list9 As List(Of Double) = Nothing, ByRef Optional list10 As List(Of Double) = Nothing,
    ByRef Optional list11 As List(Of Double) = Nothing, ByRef Optional list12 As List(Of Double) = Nothing, ByRef Optional list13 As List(Of Double) = Nothing,
    ByRef Optional list14 As List(Of Double) = Nothing, ByRef Optional list15 As List(Of Double) = Nothing, ByRef Optional list16 As List(Of Double) = Nothing)
    Dim min_num_points%
    MinNumPoints = -1

    min_num_points = date1.Count
    If list0.Count < min_num_points Then min_num_points = list0.Count
    If Not IsNothing(list1) Then
      If list1.Count < min_num_points Then min_num_points = list1.Count
    End If
    If Not IsNothing(list2) Then
      If list2.Count < min_num_points Then min_num_points = list2.Count
    End If
    If Not IsNothing(list3) Then
      If list3.Count < min_num_points Then min_num_points = list3.Count
    End If
    If Not IsNothing(list4) Then
      If list4.Count < min_num_points Then min_num_points = list4.Count
    End If
    If Not IsNothing(list5) Then
      If list5.Count < min_num_points Then min_num_points = list5.Count
    End If
    If Not IsNothing(list6) Then
      If list6.Count < min_num_points Then min_num_points = list6.Count
    End If
    If Not IsNothing(list7) Then
      If list7.Count < min_num_points Then min_num_points = list7.Count
    End If
    If Not IsNothing(list8) Then
      If list8.Count < min_num_points Then min_num_points = list8.Count
    End If
    If Not IsNothing(list9) Then
      If list9.Count < min_num_points Then min_num_points = list9.Count
    End If
    If Not IsNothing(list10) Then
      If list10.Count < min_num_points Then min_num_points = list10.Count
    End If
    If Not IsNothing(list11) Then
      If list11.Count < min_num_points Then min_num_points = list11.Count
    End If
    If Not IsNothing(list12) Then
      If list12.Count < min_num_points Then min_num_points = list12.Count
    End If
    If Not IsNothing(list13) Then
      If list13.Count < min_num_points Then min_num_points = list13.Count
    End If
    If Not IsNothing(list14) Then
      If list14.Count < min_num_points Then min_num_points = list14.Count
    End If
    If Not IsNothing(list15) Then
      If list15.Count < min_num_points Then min_num_points = list15.Count
    End If
    If Not IsNothing(list16) Then
      If list16.Count < min_num_points Then min_num_points = list16.Count
    End If

    If min_num_points > num_for_calc Then min_num_points = num_for_calc
    If min_num_points < 10 Then Exit Function

    MinNumPoints = min_num_points
  End Function
  Function DaysRisingOrFalling%(n%, open As List(Of Double), close As List(Of Double), bCheckPreviousClose As Boolean)
    ' n = maximum number of elements to check
    Dim count%, i%, ii%, num_in_list%
    DaysRisingOrFalling = 0
    count = 0
    num_in_list = close.Count
    If num_in_list < n Then
      n = num_in_list
    End If

    If bCheckPreviousClose Then
      For i = 0 To n - 2
        ii = num_in_list - 1 - i  ' ii decreases starting with the last element
        If close(ii) > close(ii - 1) Then
          If count >= 0 Then
            count += 1
          Else
            Exit For
          End If
        ElseIf close(ii) < close(ii - 1) Then
          If count <= 0 Then
            count -= 1
          Else
            Exit For
          End If
        Else
          Exit For
        End If
      Next
    Else
      For i = 0 To n - 1
        ii = num_in_list - 1 - i  ' ii decreases starting with the last element
        If close(ii) > open(ii) Then
          If count >= 0 Then
            count += 1
          Else
            Exit For
          End If
        ElseIf close(ii) < open(ii) Then
          If count <= 0 Then
            count -= 1
          Else
            Exit For
          End If
        Else
          Exit For
        End If
      Next
    End If
    DaysRisingOrFalling = count
  End Function
  Function DaysRisingOrFalling2%(j%, n%, open As List(Of Double), close As List(Of Double), bCheckPreviousClose As Boolean)
    ' j = current index in the list
    ' n = maximum number of elements to check
    Dim count%, i%, num_in_list%
    DaysRisingOrFalling2 = 0
    If j < 1 Then Exit Function
    count = 0
    num_in_list = close.Count
    If num_in_list < n Then
      n = num_in_list
    End If

    If bCheckPreviousClose Then
      If j - n + 2 < 1 Then
        n = j + 1
      End If
      For i = j To j - n + 2 Step -1
        If close(i) > close(i - 1) Then
          If count >= 0 Then
            count += 1
          Else
            Exit For
          End If
        ElseIf close(i) < close(i - 1) Then
          If count <= 0 Then
            count -= 1
          Else
            Exit For
          End If
        Else
          Exit For
        End If
      Next
    Else
      If j - n + 1 < 1 Then
        n = j
      End If
      For i = j To j - n + 1 Step -1
        If close(i) > open(i) Then
          If count >= 0 Then
            count += 1
          Else
            Exit For
          End If
        ElseIf close(i) < open(i) Then
          If count <= 0 Then
            count -= 1
          Else
            Exit For
          End If
        Else
          Exit For
        End If
      Next
    End If
    DaysRisingOrFalling2 = count
  End Function

  Function DaysRisingOrFalling3(j%, n%, open As List(Of Double), close As List(Of Double), bCheckPreviousClose As Boolean) As List(Of Integer)
    ' check from the starting index forward not backwards
    ' j = starting index in the list
    ' n = number of elements to check
    Dim count%, i%, num_in_list%
    DaysRisingOrFalling3 = Nothing
    If bCheckPreviousClose And j < 1 Then Exit Function
    DaysRisingOrFalling3 = New List(Of Integer)

    count = 0
    num_in_list = close.Count
    If num_in_list < n Then
      n = num_in_list
    End If

    If bCheckPreviousClose Then
      For i = j To n - 1
        If close(i) > close(i - 1) Then
          If count <= 0 Then
            count = 1
          Else
            count += 1
          End If
        ElseIf close(i) < close(i - 1) Then
          If count >= 0 Then
            count = -1
          Else
            count -= 1
          End If
        Else
          Exit For
        End If
        DaysRisingOrFalling3.Add(count)
      Next
    Else
      For i = j To n - 1
        If close(i) > open(i) Then
          If count <= 0 Then
            count = 1
          Else
            count += 1
          End If
        ElseIf close(i) < open(i) Then
          If count >= 0 Then
            count = -1
          Else
            count -= 1
          End If
        Else
          Exit For
        End If
        DaysRisingOrFalling3.Add(count)
      Next
    End If
  End Function
End Module

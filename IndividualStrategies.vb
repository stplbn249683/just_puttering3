Option Strict Off
Option Explicit On
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NLoptNet
Imports Skender.Stock.Indicators
Public Class PublicAccessType1
  Public quotes As IEnumerable(Of Skender.Stock.Indicators.Quote)
  Public low, high, open, close, volume, macd, histogram, mfi, rsi As New List(Of Double)
  Public vix, superTrend As New List(Of Double)
  Public date1 As New List(Of Date)
  Public isUpTrend As New List(Of Boolean)
End Class
Module IndividualStrategies
  Public pb1 As New PublicAccessType1
  Public Const NO_BUY% = 0, BUY% = 1, HOLD% = 2, STOPPED_OUT% = 3, SELL% = 4
  Public sActions$() = {"no buy", "BUY", "HOLD", "SHOULD BE STOPPED_OUT", "SELL"}
  Function number_of_opt_variables%(strategy_no%)
    number_of_opt_variables = -1
    Select Case strategy_no
      Case 1
        number_of_opt_variables = 7
      Case 2
        number_of_opt_variables = 2
      Case 3
        number_of_opt_variables = 2
    End Select
  End Function

  Sub set_bounds(StrategyNo%, ByRef solver As NLoptSolver, ByRef initialValue#())
    Select Case StrategyNo
      Case 1
        solver.SetLowerBounds(0.0, 0.0, 0.0, 0.97, 0.0, 0.0, 0.0)
        solver.SetUpperBounds(100.0, 10.0, 100.0, 0.97, 1.0, 1.0, 1.0)

        initialValue(0) = 0.0
        initialValue(1) = 0.0
        initialValue(2) = 0.0
        initialValue(3) = 0.97
        initialValue(4) = 0.0
        initialValue(5) = 0.0
        initialValue(6) = 0.0

      Case 2
        solver.SetLowerBounds(0.0, 0.0)
        solver.SetUpperBounds(60.0, 1.0)

        initialValue(0) = 0.0
        initialValue(1) = 0.0

      Case 3
        solver.SetLowerBounds(4.0, 1.0)
        solver.SetUpperBounds(20.0, 5.0)
        initialValue(0) = 7.0
        initialValue(1) = 3.0

        pb.x_type(0) = "i"
    End Select
  End Sub

  Function calculate_indicators%(StrategyNo%, max_num_points%, num_for_calc%, quotes As IEnumerable(Of Quote), connection_string$)
    calculate_indicators = -1

    pb.min_num_trades_per_year = 5.0
    pb.perc_gain_per_day_weight = 100.0
    pb.min_perc_days_in_market = 5.0
    pb.perc_days_in_market_weight = 0.2
    pb.win_rate_weight = 1.0
    pb.perc_max_drawdown_weight = 0.01
    pb.gl_ratio_weight = 0.0
    pb.perc_return_per_year_weight = 0.0
    pb.sharpe_ratio_weight = 0.0

    Dim num_for_resize = num_for_calc + 1 ' because the first point is not used in the search for buy/sell
    pb.num_for_resize = num_for_resize
    Dim max_resize = 30000
    Dim error1%
    Select Case StrategyNo
      Case 1
        pb.perc_gain_per_day_weight = 10.0
        pb.perc_max_drawdown_weight = 0.0
        pb.win_rate_weight = 0.1
        pb.gl_ratio_weight = 0.03

        Dim rsi_result = quotes.GetRsi(14)
        pb1.date1.Clear()
        Call GetRsiLists(rsi_result, pb1.date1, pb1.rsi)

        Dim mfi_result = quotes.GetMfi(14)
        pb1.date1.Clear()
        Call GetMfiLists(mfi_result, pb1.date1, pb1.mfi)

        Dim macdSignal, ema26 As New List(Of Double)
        Dim macd_result = quotes.GetMacd(12, 26, 9)
        pb1.date1.Clear()
        Call GetMacdLists(macd_result, pb1.date1, pb1.macd, macdSignal, pb1.histogram)

        Dim ema26_result = quotes.GetEma(26)
        pb1.date1.Clear()
        GetEmaLists(ema26_result, pb1.date1, ema26)
        error1 = ResizeLists(max_resize, pb1.date1, ema26, pb1.macd, macdSignal, pb1.histogram)
        If error1 < 0 Then
          MessageBox.Show("Error resizing lists")
          Exit Function
        End If

        For i = 0 To pb1.histogram.Count - 1
          pb1.histogram(i) = 100.0 * pb1.histogram(i) / ema26(i) 'Find normalized MACD histogram so the optimization parameters will be similar for different stocks
        Next

        error1 = ResizeLists(num_for_resize, pb1.date1, pb1.low, pb1.open, pb1.close, pb1.rsi, pb1.histogram, pb1.mfi)
        If error1 < 0 Then
          MessageBox.Show("Error resizing lists")
          Exit Function
        End If

      Case 2
        Dim vix_quotes As IEnumerable(Of Skender.Stock.Indicators.Quote)
        vix_quotes = GetQuotes(max_num_points, "$VIX", connection_string)
        Dim num_from_db_vix% = vix_quotes.Count
        If num_from_db_vix <= 0 Then
          MessageBox.Show("VIX ticker symbol not In database")
          Exit Function
        End If

        If num_from_db_vix <= 10 Or num_for_calc > num_from_db_vix Then
          MessageBox.Show("Not enough points for VIX ticker symbol in database")
          Exit Function
        End If

        Call GetQuoteCloseLists(vix_quotes, pb1.date1, pb1.vix)

        pb.min_num_trades_per_year = 0.2
        pb.perc_gain_per_day_weight = 100.0
        pb.perc_max_drawdown_weight = 0.0
        pb.win_rate_weight = 0.2
        pb.gl_ratio_weight = 0.0
        pb.perc_max_drawdown_weight = 1.0
        pb.min_perc_days_in_market = 5.0
        pb.perc_days_in_market_weight = 20.0

        error1 = ResizeLists(num_for_resize, pb1.date1, pb1.low, pb1.open, pb1.close, pb1.vix)
        If error1 < 0 Then
          MessageBox.Show("Error resizing lists")
          Exit Function
        End If
    End Select

    pb.num_used = pb1.date1.Count
    calculate_indicators = 0
  End Function
  Function calculate_indicators1%(x#())
    calculate_indicators1 = -1

    Select Case pb.StrategyNo
      Case 3
        'pb.min_num_trades_per_year = 4.0
        pb.min_num_trades_per_year = 0
        pb.perc_gain_per_day_weight = 0.0
        pb.perc_max_drawdown_weight = 0.0
        pb.win_rate_weight = 0.0
        'pb.win_rate_weight = 0.1
        'pb.gl_ratio_weight = 0.03
        'pb.perc_return_per_year_weight = 1.0
        pb.sharpe_ratio_weight = 1.0

        Dim max_resize = 30000

        Dim lookbackPeriods% = Math.Round(x(0))
        Dim st_result = pb1.quotes.GetSuperTrend(lookbackPeriods, x(1))
        GetSuperTrendLists1(st_result, pb1.superTrend, pb1.isUpTrend)

        Dim min_num_points%
        min_num_points = MinNumPoints(pb.num_for_resize, pb1.date1, pb1.low, pb1.close, pb1.superTrend)
        If min_num_points < 0 Then
          MessageBox.Show("Error in MinNumPoints")
          Exit Function
        End If

        pb.num_used = min_num_points
    End Select
    calculate_indicators1 = 0
  End Function
  Function results_file_heading$(StrategyNo%)
    Dim s1$ = ""
    Select Case StrategyNo
      Case 1
        s1 = ",RSI,MACD histogram,MFI"
      Case 2
        s1 = ",VIX"
      Case 3
        s1 = ",SuperTrend, Is UpTrend"
    End Select
    results_file_heading = s1
  End Function

  Function results_file_text(j%, StrategyNo%)
    Dim s1$ = ""
    Select Case StrategyNo
      Case 1
        s1 = "," & pb1.rsi(j).ToString("0.000") & "," & pb1.histogram(j).ToString("0.000") & "," & pb1.mfi(j).ToString("0.000")
      Case 2
        s1 = "," & pb1.vix(j).ToString("0.000")
      Case 3
        Dim j_st%
        j_st = j + pb1.superTrend.Count - pb.num_used
        s1 = "," & pb1.superTrend(j_st).ToString("0.000") & "," & pb1.isUpTrend(j_st).ToString()
    End Select
    results_file_text = s1
  End Function
  Function find_action%(j%, StrategyNo%, ByRef x#(), bBought As Boolean, price_bought#, ByRef price_sold#)
    Dim action1%
    action1 = NO_BUY

    Select Case StrategyNo
      Case 1
        ' x(0),x(1),x(2) = RSI, MACD histogram, MFI for buy
        ' x(3) = stop fraction
        ' x(4),x(5),x(6) = RSI, MACD histogram, MFI multipliers for sell

        If Not bBought Then
          action1 = NO_BUY
          If pb1.rsi(j) > x(0) And pb1.histogram(j) > x(1) And pb1.mfi(j) > x(2) Then
            action1 = BUY
          End If
        Else
          action1 = HOLD

          If pb1.close(j) < x(3) * price_bought Then
            action1 = STOPPED_OUT
            price_sold = x(3) * price_bought
          Else
            If pb1.rsi(j) < x(0) * x(4) Or pb1.histogram(j) < x(1) * x(5) Or pb1.mfi(j) < x(2) * x(6) Then
              action1 = SELL
              price_sold = pb1.close(j) ' price from close the same day
            End If
          End If
        End If

      Case 2
        ' x(0),x(1) = VIX for buy, VIX multiplier for sell
        If Not bBought Then
          action1 = NO_BUY
          If pb1.vix(j) > x(0) Then
            action1 = BUY
          End If
        Else
          action1 = HOLD

          'If pb1.close(j) < x(4) * price_bought Then
          '  action1 = STOPPED_OUT
          '  price_sold = x(4) * price_bought
          'Else
          If pb1.vix(j) <= x(0) * x(1) Then
            action1 = SELL
            price_sold = pb1.close(j) ' price from close the same day
          End If
          'End If
        End If

      Case 3
        ' x(0),x(1) = look back periods, multiplier

        Dim j_st = j + pb1.superTrend.Count - pb.num_used
        Dim j_cl = j + pb1.close.Count - pb.num_used
        Dim j_l = j + pb1.low.Count - pb.num_used

        If Not bBought Then
          action1 = NO_BUY
          If pb1.isUpTrend(j_st) Then
            action1 = BUY
          End If
        Else
          action1 = HOLD
          If Not pb1.isUpTrend(j_st) Then
            action1 = SELL
            price_sold = pb1.close(j_cl) ' price from close the same day
            pb.sell_date_index = j
          End If

          'If pb1.close(j_cl) < 0.96 * price_bought Or pb1.low(j_l) < 0.96 * price_bought Then
          '  action1 = STOPPED_OUT
          '  price_sold = 0.96 * price_bought
          '  pb.sell_date_index = j
          'Else
          '  If Not pb1.isUpTrend(j_st) Then
          '    action1 = SELL
          '    price_sold = pb1.close(j_cl) ' price from close the same day
          '    pb.sell_date_index = j
          '  End If
          'End If
        End If
    End Select
    find_action = action1
  End Function
End Module

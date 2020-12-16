Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class BTSTValueInvestingStrategyRule
    Inherits StrategyRule

    Public LastTick As Payload = Nothing
    Public EntryATR As Decimal = Decimal.MinValue

    Private _eodPayload As Dictionary(Of Date, Payload)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _rainbowPayload As Dictionary(Of Date, Indicator.RainbowMA)
    Private _lastTrade As Trade = Nothing

    Private ReadOnly _cashBrokerage As Decimal = 0.11 / 100
    Private ReadOnly _futBrokerage As Decimal = 0.01 / 100

    Private ReadOnly _controller As Boolean

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal entities As RuleEntities,
                   ByVal canceller As CancellationTokenSource,
                   ByVal controller As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, entities, canceller)

        If controller = 1 Then
            _controller = True
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
        If _controller Then
            _eodPayload = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_POSITIONAL, _tradingSymbol, _tradingDate.AddDays(200), _tradingDate)
            If _eodPayload IsNot Nothing AndAlso _eodPayload.Count > 100 Then
                Indicator.ATR.CalculateATR(14, _eodPayload, _atrPayload)
                Indicator.RainbowMovingAverage.CalculateRainbowMovingAverage(2, _eodPayload, _rainbowPayload)
            Else
                Throw New ApplicationException("Can not trade")
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If (Me.ForceTakeTrade OrElse Me.ContractRolloverForceEntry) AndAlso Not _controller Then
            Dim quantity As Integer = Me.LotSize

            Dim avgPrice As Decimal = currentTick.Open
            Dim entryTrades As List(Of Trade) = GetEntryTrades()
            If entryTrades IsNot Nothing AndAlso entryTrades.Count Then
                Dim totalValue As Decimal = entryTrades.Sum(Function(x)
                                                                Return x.EntryPrice * x.Quantity
                                                            End Function)
                Dim totalQty As Decimal = entryTrades.Sum(Function(x)
                                                              Return x.Quantity
                                                          End Function)
                totalValue += currentTick.Open * quantity
                totalQty += quantity
                avgPrice = totalValue / totalQty
            End If

            If Me.ContractRolloverForceEntry AndAlso _lastTrade IsNot Nothing Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 100000,
                                .Target = .EntryPrice + 100000,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = avgPrice,
                                .Supporting2 = _lastTrade.Supporting2,
                                .Supporting3 = _lastTrade.Supporting3,
                                .Supporting4 = _lastTrade.Supporting4
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.ContractRolloverForceEntry = False
            ElseIf Me.ForceTakeTrade Then
                Dim parameter As PlaceOrderParameters = Nothing
                parameter = New PlaceOrderParameters With {
                                .EntryPrice = currentTick.Open,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = .EntryPrice - 100000,
                                .Target = .EntryPrice + 100000,
                                .Buffer = 0,
                                .SignalCandle = currentMinuteCandle,
                                .OrderType = Trade.TypeOfOrder.Market,
                                .Supporting1 = avgPrice,
                                .Supporting2 = Me.EntryATR,
                                .Supporting3 = System.Guid.NewGuid.ToString(),
                                .Supporting4 = currentTick.Open * quantity
                            }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                Me.ForceTakeTrade = False
            End If
        ElseIf _controller Then
            Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
            If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso _eodPayload IsNot Nothing AndAlso
                _eodPayload.ContainsKey(_tradingDate.Date) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.CNC) AndAlso
                currentMinuteCandle.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
                If currentMinuteCandle.PayloadDate = _signalPayload.LastOrDefault.Key Then
                    If Not Me.IsActiveTrade() Then
                        Dim currentDayCandle As Payload = _eodPayload(_tradingDate.Date)
                        If currentDayCandle.CandleColor = Color.Green Then
                            Dim rainbow As Indicator.RainbowMA = _rainbowPayload(currentDayCandle.PayloadDate)
                            If currentDayCandle.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                                If IsValidRainbow(currentDayCandle) Then
                                    Me.AnotherPairInstrument.ForceTakeTrade = True
                                    CType(Me.AnotherPairInstrument, BTSTValueInvestingStrategyRule).EntryATR = _atrPayload(_tradingDate.Date)
                                End If
                            End If
                        End If
                    Else
                        Dim entryTrades As List(Of Trade) = GetEntryTrades()
                        If entryTrades IsNot Nothing AndAlso entryTrades.Count Then
                            Dim lastTrade As Trade = entryTrades.OrderByDescending(Function(x)
                                                                                       Return x.EntryTime
                                                                                   End Function).LastOrDefault
                            If _tradingDate.Date >= lastTrade.EntryTime.Date.AddDays(7) Then
                                Dim cashTradeCount As Integer = entryTrades.Sum(Function(x)
                                                                                    If Not x.TradingSymbol.EndsWith("FUT") Then
                                                                                        Return 1
                                                                                    Else
                                                                                        Return 0
                                                                                    End If
                                                                                End Function)
                                Dim initialInvestment As Decimal = lastTrade.Supporting4
                                Dim desireIncrease As Decimal = initialInvestment * 1 / 100
                                Dim desireValue As Decimal = initialInvestment + (desireIncrease * (cashTradeCount + 1))

                                Dim cashTick As Payload = Me.LastTick
                                Dim futTick As Payload = CType(Me.AnotherPairInstrument, BTSTValueInvestingStrategyRule).LastTick
                                Dim totalCashQuantity As Integer = entryTrades.Sum(Function(x)
                                                                                       If Not x.TradingSymbol.EndsWith("FUT") Then
                                                                                           Return x.Quantity
                                                                                       Else
                                                                                           Return 0
                                                                                       End If
                                                                                   End Function)

                                Dim totalFutQuantity As Integer = entryTrades.Sum(Function(x)
                                                                                      If x.TradingSymbol.EndsWith("FUT") Then
                                                                                          Return x.Quantity
                                                                                      Else
                                                                                          Return 0
                                                                                      End If
                                                                                  End Function)
                                Dim bkrgCsh As Decimal = totalCashQuantity * cashTick.Close * _cashBrokerage
                                Dim bkrgFut As Decimal = totalFutQuantity * futTick.Close * _futBrokerage
                                Dim totalPortfolioValueBeforeRebalancing As Decimal = (totalCashQuantity * cashTick.Close) + bkrgCsh + (totalFutQuantity * futTick.Close) + bkrgFut
                                Dim amountToInvest As Decimal = desireValue - totalPortfolioValueBeforeRebalancing
                                If amountToInvest > 0 Then
                                    Dim cashSharesToBuy As Integer = Math.Ceiling(amountToInvest / cashTick.Close)
                                    If cashSharesToBuy > 0 Then
                                        Dim totalValue As Decimal = entryTrades.Sum(Function(x)
                                                                                        Return x.EntryPrice * x.Quantity
                                                                                    End Function)
                                        Dim totalQty As Decimal = entryTrades.Sum(Function(x)
                                                                                      Return x.Quantity
                                                                                  End Function)
                                        totalValue += currentTick.Open * cashSharesToBuy
                                        totalQty += cashSharesToBuy
                                        Dim avgPrice As Decimal = totalValue / totalQty

                                        Dim parameter As PlaceOrderParameters = Nothing
                                        parameter = New PlaceOrderParameters With {
                                                        .EntryPrice = currentTick.Open,
                                                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                        .Quantity = cashSharesToBuy,
                                                        .Stoploss = .EntryPrice - 100000,
                                                        .Target = .EntryPrice + 100000,
                                                        .Buffer = 0,
                                                        .SignalCandle = currentMinuteCandle,
                                                        .OrderType = Trade.TypeOfOrder.Market,
                                                        .Supporting1 = avgPrice,
                                                        .Supporting2 = lastTrade.Supporting2,
                                                        .Supporting3 = lastTrade.Supporting3,
                                                        .Supporting4 = lastTrade.Supporting4
                                                    }

                                        For Each runningTrade In entryTrades
                                            runningTrade.UpdateTrade(Supporting1:=avgPrice)
                                        Next
                                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If Not _controller Then
                If Me.ForceCancelTrade Then
                    ret = New Tuple(Of Boolean, String)(True, "Target Hit")
                    Me.ForceCancelTrade = False
                ElseIf Me.ContractRollover OrElse Me.BlankDayExit Then
                    Dim currentMinutePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                    If currentMinutePayload.PayloadDate = _signalPayload.LastOrDefault.Key Then
                        ret = New Tuple(Of Boolean, String)(True, If(Me.ContractRollover, "Contract Rollover Exit", "Blank Day Exit"))
                        Me.ForceCancellationDone = True
                        If Me.BlankDayExit Then
                            Me.AnotherPairInstrument.ForceCancelTrade = True
                        End If
                        _lastTrade = currentTrade
                    End If
                End If
            ElseIf _controller Then
                If Me.ForceCancelTrade Then
                    ret = New Tuple(Of Boolean, String)(True, "Blank Day Cash Exit")
                    Me.ForceCancelTrade = False
                Else
                    Dim avgPrice As Decimal = currentTrade.Supporting1
                    If Me.LastTick.Close >= avgPrice + avgPrice * 1 / 100 Then
                        ret = New Tuple(Of Boolean, String)(True, "Target Hit")
                        Me.AnotherPairInstrument.ForceCancelTrade = True
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Me.LastTick = currentTick
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Private Function IsValidRainbow(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        For Each runningPayload In _eodPayload.OrderByDescending(Function(x)
                                                                     Return x.Key
                                                                 End Function)
            If runningPayload.Key <= candle.PreviousCandlePayload.PayloadDate Then
                Dim rainbow As Indicator.RainbowMA = _rainbowPayload(runningPayload.Key)
                If runningPayload.Value.Close > Math.Max(rainbow.SMA1, Math.Max(rainbow.SMA2, Math.Max(rainbow.SMA3, Math.Max(rainbow.SMA4, Math.Max(rainbow.SMA5, Math.Max(rainbow.SMA6, Math.Max(rainbow.SMA7, Math.Max(rainbow.SMA8, Math.Max(rainbow.SMA9, rainbow.SMA10))))))))) Then
                    Exit For
                ElseIf runningPayload.Value.Close < Math.Min(rainbow.SMA1, Math.Min(rainbow.SMA2, Math.Min(rainbow.SMA3, Math.Min(rainbow.SMA4, Math.Min(rainbow.SMA5, Math.Min(rainbow.SMA6, Math.Min(rainbow.SMA7, Math.Min(rainbow.SMA8, Math.Min(rainbow.SMA9, rainbow.SMA10))))))))) Then
                    If runningPayload.Value.CandleColor = Color.Red Then
                        ret = True
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function

    Private Function IsActiveTrade() As Boolean
        Return (_parentStrategy.IsTradeActive(Me.LastTick, Trade.TypeOfTrade.CNC) OrElse
            _parentStrategy.IsTradeActive(CType(Me.AnotherPairInstrument, BTSTValueInvestingStrategyRule).LastTick, Trade.TypeOfTrade.CNC))
    End Function

    Private Function GetEntryTrades() As List(Of Trade)
        Dim ret As List(Of Trade) = Nothing
        Dim myTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(Me.LastTick, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        Dim myFutTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(CType(Me.AnotherPairInstrument, BTSTValueInvestingStrategyRule).LastTick, Trade.TypeOfTrade.CNC, Trade.TradeExecutionStatus.Inprogress)
        If myTrades IsNot Nothing AndAlso myFutTrades IsNot Nothing Then
            ret = New List(Of Trade)
            ret.AddRange(myTrades)
            ret.AddRange(myFutTrades)
        Else
            If myTrades IsNot Nothing Then ret = myTrades
            If myFutTrades IsNot Nothing Then ret = myFutTrades
        End If
        Return ret
    End Function

End Class

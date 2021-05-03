Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper

Public Class ATFNiftySwingTradingWithMatingale
    Inherits StrategyRule

    Private ReadOnly _buyPrice As Decimal
    Private ReadOnly _sellPrice As Decimal
    Private ReadOnly _lastTradingDayOfThisContract As Boolean
    Private ReadOnly _eodExitTime As Date
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal buyPrice As Decimal,
                   ByVal sellPrice As Decimal,
                   ByVal lastTradingDayOfThisContract As Boolean)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _buyPrice = buyPrice
        _sellPrice = sellPrice
        _lastTradingDayOfThisContract = lastTradingDayOfThisContract
        _eodExitTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.EODExitTime.Hours, _parentStrategy.EODExitTime.Minutes, _parentStrategy.EODExitTime.Seconds)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso (Not _lastTradingDayOfThisContract OrElse
            currentTick.PayloadDate < _eodExitTime) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, _parentStrategy.TradeType) Then
            Dim entryDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
            Dim iteration As Integer = 1
            Dim previousLoss As Decimal = 0
            If currentTick.Open >= _buyPrice Then
                entryDirection = Trade.TradeExecutionDirection.Buy
            ElseIf currentTick.Open <= _sellPrice Then
                entryDirection = Trade.TradeExecutionDirection.Sell
            End If
            If entryDirection <> Trade.TradeExecutionDirection.None Then
                Dim lastTrade As Trade = _parentStrategy.GetLastEntryTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.CNC)
                If lastTrade IsNot Nothing AndAlso lastTrade.ExitRemark.ToUpper = "STOPLOSS" Then
                    iteration = Val(lastTrade.Supporting4) + 1
                    previousLoss = Val(lastTrade.Supporting5) + lastTrade.PLAfterBrokerage
                End If
                If lastTrade IsNot Nothing AndAlso lastTrade.EntryDirection = entryDirection Then
                    If lastTrade.ExitRemark.ToUpper = "TARGET" Then
                        If lastTrade.Supporting2 = _buyPrice AndAlso lastTrade.Supporting3 = _sellPrice Then
                            entryDirection = Trade.TradeExecutionDirection.None
                        End If
                    End If
                End If
            End If
            If entryDirection = Trade.TradeExecutionDirection.Buy Then
                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = Me.LotSize * Math.Pow(2, iteration - 1),
                            .Stoploss = .EntryPrice - 1000000,
                            .Target = .EntryPrice + 1000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandle,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = currentMinuteCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                            .Supporting2 = _buyPrice,
                            .Supporting3 = _sellPrice,
                            .Supporting4 = iteration,
                            .Supporting5 = previousLoss
                        }
            ElseIf entryDirection = Trade.TradeExecutionDirection.Sell Then
                parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = Me.LotSize * Math.Pow(2, iteration - 1),
                            .Stoploss = .EntryPrice + 1000000,
                            .Target = .EntryPrice - 1000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandle,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = currentMinuteCandle.PayloadDate.ToString("dd-MMM-yyyy HH:mm:ss"),
                            .Supporting2 = _buyPrice,
                            .Supporting3 = _sellPrice,
                            .Supporting4 = iteration,
                            .Supporting5 = previousLoss
                        }
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open <= _sellPrice Then
                    ret = New Tuple(Of Boolean, String)(True, "Stoploss")
                ElseIf currentTick.Open >= currentTrade.EntryPrice + 100 Then
                    If currentTrade.Supporting4 = 1 Then
                        ret = New Tuple(Of Boolean, String)(True, "Target")
                    Else
                        If currentTrade.PLAfterBrokerage + Val(currentTrade.Supporting5) >= Math.Abs(Val(currentTrade.Supporting5)) / (Math.Pow(2, Val(currentTrade.Supporting4) - 1) - 1) Then
                            ret = New Tuple(Of Boolean, String)(True, "Target")
                        End If
                    End If
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open >= _buyPrice Then
                    ret = New Tuple(Of Boolean, String)(True, "Stoploss")
                ElseIf currentTick.Open <= currentTrade.EntryPrice - 100 Then
                    If currentTrade.Supporting4 = 1 Then
                        ret = New Tuple(Of Boolean, String)(True, "Target")
                    Else
                        If currentTrade.PLAfterBrokerage + Val(currentTrade.Supporting5) >= Math.Abs(Val(currentTrade.Supporting5)) / (Math.Pow(2, Val(currentTrade.Supporting4) - 1) - 1) Then
                            ret = New Tuple(Of Boolean, String)(True, "Target")
                        End If
                    End If
                End If
            End If
            If _lastTradingDayOfThisContract AndAlso currentTick.PayloadDate >= _eodExitTime Then
                ret = New Tuple(Of Boolean, String)(True, "EOC Exit")
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function
End Class
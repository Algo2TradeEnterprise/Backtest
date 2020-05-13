Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class XMinVWAPStrategy
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetPercentage As Decimal
        Public MaxNumberOfIteration As Decimal
        Public MinimumInvestmentPerStock As Integer
    End Class
#End Region

    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _vwapPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Private ReadOnly _firstSignalCandleTime As Date
    Private ReadOnly _lastSignalCandleTime As Date
    Private ReadOnly _lastTradeEntryTime As Date
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _lastTradeEntryTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.LastTradeEntryTime.Hours, _parentStrategy.LastTradeEntryTime.Minutes, _parentStrategy.LastTradeEntryTime.Seconds)
        _firstSignalCandleTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 9, 30, 0)
        _lastSignalCandleTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 10, 0, 0)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.SMA.CalculateSMA(10, Payload.PayloadFields.Volume, _signalPayload, _smaPayload)
        Indicator.VWAP.CalculateVWAP(_signalPayload, _vwapPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeOpen(currentTick, _parentStrategy.TradeType) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso currentMinuteCandlePayload.PayloadDate <= _lastTradeEntryTime Then

            If currentMinuteCandlePayload.PayloadDate > _lastSignalCandleTime.AddMinutes(_parentStrategy.SignalTimeFrame) Then
                EligibleToTakeTrade = False
            End If

            Dim signalReceivedForEntry As Tuple(Of Boolean, Integer) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalReceivedForEntry IsNot Nothing AndAlso signalReceivedForEntry.Item1 Then
                Dim investment As Decimal = Math.Pow(2, signalReceivedForEntry.Item2) * _userInputs.MinimumInvestmentPerStock
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromInvestment(Me.LotSize, investment, currentTick.Open, Trade.TypeOfStock.Cash, False)
                If _direction = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice - 1000000000,
                            .Target = .EntryPrice + 1000000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandlePayload,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = signalReceivedForEntry.Item2
                        }
                ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                            .EntryPrice = currentTick.Open,
                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                            .Quantity = quantity,
                            .Stoploss = .EntryPrice + 1000000000,
                            .Target = .EntryPrice - 1000000000,
                            .Buffer = 0,
                            .SignalCandle = currentMinuteCandlePayload,
                            .OrderType = Trade.TypeOfOrder.Market,
                            .Supporting1 = signalReceivedForEntry.Item2
                        }
                End If
                Dim totalCapitalUsedWithoutMargin As Decimal = 0
                Dim totalQuantity As Decimal = 0
                Dim openActiveTrades As List(Of Trade) = _parentStrategy.GetOpenActiveTrades(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                    For Each runningTrade In openActiveTrades
                        If runningTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                            totalCapitalUsedWithoutMargin += runningTrade.EntryPrice * runningTrade.Quantity
                            totalQuantity += runningTrade.Quantity
                        End If
                    Next
                End If
                totalCapitalUsedWithoutMargin += parameter.EntryPrice * parameter.Quantity
                totalQuantity += parameter.Quantity
                Dim averageTradePrice As Decimal = ConvertFloorCeling(totalCapitalUsedWithoutMargin / totalQuantity, _parentStrategy.TickSize, RoundOfType.Floor)
                If openActiveTrades IsNot Nothing AndAlso openActiveTrades.Count > 0 Then
                    For Each runningTrade In openActiveTrades
                        runningTrade.UpdateTrade(Supporting2:=averageTradePrice)
                    Next
                End If
                parameter.Supporting2 = averageTradePrice
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
            Dim averagePrice As Decimal = currentTrade.Supporting2
            Dim targetPoint As Decimal = ConvertFloorCeling(averagePrice * _userInputs.TargetPercentage / 100, _parentStrategy.TickSize, RoundOfType.Floor)
            If _direction = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open >= averagePrice + targetPoint Then
                    ret = New Tuple(Of Boolean, String)(True, "Target")
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open <= averagePrice - targetPoint Then
                    ret = New Tuple(Of Boolean, String)(True, "Target")
                End If
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

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Integer)
        Dim ret As Tuple(Of Boolean, Integer) = Nothing
        If _direction = Trade.TradeExecutionDirection.None Then
            If candle.Close < candle.Open Then
                If candle.Close < candle.Open * 0.99 Then
                    If candle.Volume > _smaPayload(candle.PayloadDate) * 2 Then
                        If candle.Close < _vwapPayload(candle.PayloadDate) AndAlso
                            candle.PreviousCandlePayload.Close > _vwapPayload(candle.PreviousCandlePayload.PayloadDate) Then
                            If candle.PayloadDate >= _firstSignalCandleTime AndAlso candle.PayloadDate <= _lastSignalCandleTime Then
                                _direction = Trade.TradeExecutionDirection.Sell
                            End If
                        End If
                    End If
                End If
            ElseIf candle.Close > candle.Open Then
                If candle.Close > candle.Open * 1.01 Then
                    If candle.Volume > _smaPayload(candle.PayloadDate) * 2 Then
                        If candle.Close > _vwapPayload(candle.PayloadDate) AndAlso
                            candle.PreviousCandlePayload.Close < _vwapPayload(candle.PreviousCandlePayload.PayloadDate) Then
                            If candle.PayloadDate >= _firstSignalCandleTime AndAlso candle.PayloadDate <= _lastSignalCandleTime Then
                                _direction = Trade.TradeExecutionDirection.Buy
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If _direction <> Trade.TradeExecutionDirection.None Then
            Dim lastExecutedTrade As Trade = GetLastOrder(currentTick)
            If lastExecutedTrade Is Nothing Then
                ret = New Tuple(Of Boolean, Integer)(True, 0)
            Else
                If lastExecutedTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                    Dim lastTradeEntryPrice As Decimal = lastExecutedTrade.EntryPrice
                    Dim lastTradeIteration As Integer = lastExecutedTrade.Supporting1
                    If lastTradeIteration + 1 < _userInputs.MaxNumberOfIteration Then
                        Dim reEntryPoint As Decimal = ConvertFloorCeling(lastTradeEntryPrice * 1 / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                        If _direction = Trade.TradeExecutionDirection.Buy Then
                            If currentTick.Open <= lastTradeEntryPrice - reEntryPoint Then
                                ret = New Tuple(Of Boolean, Integer)(True, lastTradeIteration + 1)
                            End If
                        ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                            If currentTick.Open >= lastTradeEntryPrice + reEntryPoint Then
                                ret = New Tuple(Of Boolean, Integer)(True, lastTradeIteration + 1)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetLastOrder(ByVal currentPayload As Payload) As Trade
        Dim ret As Trade = Nothing
        If currentPayload IsNot Nothing Then
            Dim lastEntryOrder As Trade = Me._parentStrategy.GetLastEntryTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
            Dim lastClosedOrder As Trade = Me._parentStrategy.GetLastExitTradeOfTheStock(currentPayload, Me._parentStrategy.TradeType)
            If lastEntryOrder IsNot Nothing Then ret = lastEntryOrder
            If lastClosedOrder IsNot Nothing AndAlso lastClosedOrder.ExitTime > lastEntryOrder.EntryTime Then
                ret = lastClosedOrder
            End If
        End If
        Return ret
    End Function
End Class
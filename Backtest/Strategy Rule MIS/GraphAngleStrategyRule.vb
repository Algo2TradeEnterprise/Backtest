Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers

Public Class GraphAngleStrategyRule
    Inherits StrategyRule

    Private ReadOnly _tradeStartTime As Date
    Private ReadOnly _direction As Trade.TradeExecutionDirection

    Private _level As Decimal
    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal tradeStartTime As Date,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)

        _tradeStartTime = tradeStartTime
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _tradeStartTime.Hour, _tradeStartTime.Minute, _tradeStartTime.Second)
        If _direction = Trade.TradeExecutionDirection.Buy Then
            _level = _signalPayload.Max(Function(x)
                                            If x.Key.Date = _tradingDate AndAlso x.Key <= tradeStartTime Then
                                                Return x.Value.High
                                            Else
                                                Return Decimal.MinValue
                                            End If
                                        End Function)
        ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
            _level = _signalPayload.Min(Function(x)
                                            If x.Key.Date = _tradingDate AndAlso x.Key <= tradeStartTime Then
                                                Return x.Value.Low
                                            Else
                                                Return Decimal.MaxValue
                                            End If
                                        End Function)
        End If

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _tradeStartTime.Hour, _tradeStartTime.Minute, _tradeStartTime.Second)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, _parentStrategy.TradeType) AndAlso
            currentMinuteCandlePayload.PayloadDate > tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, String) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                If _direction = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = Me.LotSize,
                                .Stoploss = .EntryPrice - 10000000,
                                .Target = .EntryPrice + 10000000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signal.Item3,
                                .Supporting2 = _tradeStartTime.ToString("HH:mm:ss")
                            }
                ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = signal.Item2,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = Me.LotSize,
                                .Stoploss = .EntryPrice + 10000000,
                                .Target = .EntryPrice - 10000000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signal.Item3,
                                .Supporting2 = _tradeStartTime.ToString("HH:mm:ss")
                            }
                End If
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
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, String) = GetSignalForEntry(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If signal.Item2 <> currentTrade.EntryPrice Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim entryCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTrade.EntryTime, _signalPayload))
            If entryCandle.PayloadDate = currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                Dim triggerPrice As Decimal = Decimal.MinValue
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    triggerPrice = _fractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) - currentTrade.StoplossBuffer
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    triggerPrice = _fractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate) + currentTrade.StoplossBuffer
                End If

                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, Math.Abs(triggerPrice - currentTrade.EntryPrice))
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetSignalForEntry(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, String)
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(_level, RoundOfType.Floor)
            If _direction = Trade.TradeExecutionDirection.Buy Then
                If _fractalHighPayload(candle.PayloadDate) < _fractalHighPayload(candle.PreviousCandlePayload.PayloadDate) AndAlso
                    _fractalHighPayload(candle.PayloadDate) < _level Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, _fractalHighPayload(candle.PayloadDate) + buffer, "Fractal")
                Else
                    ret = New Tuple(Of Boolean, Decimal, String)(True, _level + buffer, "HH")
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                If _fractalLowPayload(candle.PayloadDate) > _fractalLowPayload(candle.PreviousCandlePayload.PayloadDate) AndAlso
                    _fractalLowPayload(candle.PayloadDate) > _level Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, _fractalLowPayload(candle.PayloadDate) - buffer, "Fractal")
                Else
                    ret = New Tuple(Of Boolean, Decimal, String)(True, _level - buffer, "LL")
                End If
            End If
        End If
        Return ret
    End Function
End Class
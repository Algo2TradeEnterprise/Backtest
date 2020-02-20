Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FavourableFractalBreakoutStrategyRule2
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPercentagePerTrade As Decimal
        Public MaxProfitPercentagePerTrade As Decimal
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)

    Private _higherSignalCandle As Payload = Nothing
    Private _lowerSignalCandle As Payload = Nothing

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If Not (lastExecutedOrder IsNot Nothing AndAlso lastExecutedOrder.SignalCandle.PayloadDate = signal.Item4.PayloadDate) Then
                    signalCandle = signal.Item4
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entry As Decimal = signal.Item2 + buffer
                    Dim stoploss As Decimal = ConvertFloorCeling(entry - entry * _userInputs.MaxLossPercentagePerTrade / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim target As Decimal = ConvertFloorCeling(entry + entry * _userInputs.MaxProfitPercentagePerTrade / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                    Dim entry As Decimal = signal.Item2 - buffer
                    Dim stoploss As Decimal = ConvertFloorCeling(entry + entry * _userInputs.MaxLossPercentagePerTrade / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim target As Decimal = ConvertFloorCeling(entry - entry * _userInputs.MaxProfitPercentagePerTrade / 100, _parentStrategy.TickSize, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entry,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = Me.LotSize,
                                    .Stoploss = stoploss,
                                    .Target = target,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim currentFractal As Decimal = _fractalHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                    Dim signalFractal As Decimal = _fractalHighPayload(signalCandle.PayloadDate)
                    If currentFractal <> signalFractal Then
                        If _higherSignalCandle.PayloadDate = signalCandle.PayloadDate Then _higherSignalCandle = Nothing
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim currentFractal As Decimal = _fractalLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate)
                    Dim signalFractal As Decimal = _fractalLowPayload(signalCandle.PayloadDate)
                    If currentFractal <> signalFractal Then
                        If _lowerSignalCandle.PayloadDate = signalCandle.PayloadDate Then _higherSignalCandle = Nothing
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
                If ret Is Nothing Then
                    Dim signal As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                    If signal IsNot Nothing AndAlso signal.Item1 AndAlso signal.Item3 <> currentTrade.EntryDirection Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)
        Dim ret As Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            If _higherSignalCandle IsNot Nothing AndAlso _fractalHighPayload(candle.PayloadDate) <> _fractalHighPayload(_higherSignalCandle.PayloadDate) Then
                _higherSignalCandle = Nothing
            End If
            If _lowerSignalCandle IsNot Nothing AndAlso _fractalLowPayload(candle.PayloadDate) <> _fractalLowPayload(_lowerSignalCandle.PayloadDate) Then
                _lowerSignalCandle = Nothing
            End If

            If _fractalHighPayload(candle.PayloadDate) < _fractalHighPayload(candle.PreviousCandlePayload.PayloadDate) Then
                _higherSignalCandle = candle
            End If
            If _fractalLowPayload(candle.PayloadDate) > _fractalLowPayload(candle.PreviousCandlePayload.PayloadDate) Then
                _lowerSignalCandle = candle
            End If

            If _higherSignalCandle IsNot Nothing AndAlso _lowerSignalCandle IsNot Nothing Then
                Dim buyPrice As Decimal = _fractalHighPayload(_higherSignalCandle.PayloadDate)
                Dim sellPrice As Decimal = _fractalLowPayload(_lowerSignalCandle.PayloadDate)
                Dim middle As Decimal = (buyPrice + sellPrice) / 2
                Dim range As Decimal = (buyPrice - middle) * 40 / 100
                If currentTick.Open > middle + range Then
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalHighPayload(_higherSignalCandle.PayloadDate), Trade.TradeExecutionDirection.Buy, _higherSignalCandle)
                ElseIf currentTick.Open > middle + range Then
                    ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalLowPayload(_lowerSignalCandle.PayloadDate), Trade.TradeExecutionDirection.Sell, _lowerSignalCandle)
                End If
            ElseIf _higherSignalCandle IsNot Nothing Then
                ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalHighPayload(_higherSignalCandle.PayloadDate), Trade.TradeExecutionDirection.Buy, _higherSignalCandle)
            ElseIf _lowerSignalCandle IsNot Nothing Then
                ret = New Tuple(Of Boolean, Decimal, Trade.TradeExecutionDirection, Payload)(True, _fractalLowPayload(_lowerSignalCandle.PayloadDate), Trade.TradeExecutionDirection.Sell, _lowerSignalCandle)
            End If
        End If
        Return ret
    End Function

End Class
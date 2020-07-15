Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HKReversalSingleTradeStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _slPoint As Decimal = Decimal.MinValue
    Private _targetPoint As Decimal = Decimal.MinValue
    Private _quantity As Integer = Integer.MinValue
    Private _previousDayHighestATR As Decimal = Decimal.MinValue
    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)

        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim firstCandle As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                        firstCandle = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
            If firstCandle IsNot Nothing AndAlso firstCandle.PreviousCandlePayload IsNot Nothing Then
                _previousDayHighestATR = _atrPayload.Max(Function(x)
                                                             If x.Key.Date >= firstCandle.PreviousCandlePayload.PayloadDate.Date AndAlso
                                                             x.Key.Date < firstCandle.PayloadDate.Date Then
                                                                 Return x.Value
                                                             Else
                                                                 Return Decimal.MinValue
                                                             End If
                                                         End Function)
            End If
        End If
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3
                    _slPoint = ConvertFloorCeling(GetAverageHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                    _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signal.Item2, signal.Item2 - _slPoint, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)
                    _targetPoint = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signal.Item2, _quantity, Math.Abs(_userInputs.MaxLossPerTrade * _userInputs.TargetMultiplier), Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - signal.Item2
                ElseIf lastExecutedOrder IsNot Nothing Then
                    If signal.Item3.PayloadDate <> lastExecutedOrder.SignalCandle.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim slPoint As Decimal = _slPoint
                Dim targetPoint As Decimal = _targetPoint
                Dim quantity As Integer = _quantity
                If lastExecutedOrder IsNot Nothing Then
                    quantity = lastExecutedOrder.Quantity * 2
                End If
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item2

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice - slPoint,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _previousDayHighestATR
                                }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item2

                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = .EntryPrice + slPoint,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = _previousDayHighestATR
                                }
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        ''If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
        ''    Dim triggerPrice As Decimal = Decimal.MinValue
        ''    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Math.Abs(_userInputs.MaxLossPerTrade), currentTrade.EntryDirection, Trade.TypeOfStock.Cash)
        ''    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
        ''        If currentTick.Open >= target Then
        ''            triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Buy, Me.LotSize, _parentStrategy.StockType)
        ''        End If
        ''    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
        ''        If currentTick.Open <= target Then
        ''            triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Sell, Me.LotSize, _parentStrategy.StockType)
        ''        End If
        ''    End If
        ''    If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
        ''        Dim slPoint As Decimal = currentTrade.SLRemark
        ''        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("({0})Breakeven at {1}", slPoint, currentTick.PayloadDate.ToString("HH:mm:ss")))
        ''    End If
        ''End If
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            If _direction = Trade.TradeExecutionDirection.Buy AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If buyLevel = Math.Round(hkCandle.High, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(hkCandle.High, RoundOfType.Floor)
                    buyLevel = buyLevel + buffer
                End If
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, hkCandle, Trade.TradeExecutionDirection.Buy)
            ElseIf _direction = Trade.TradeExecutionDirection.Sell AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If sellLevel = Math.Round(hkCandle.Low, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(hkCandle.Low, RoundOfType.Floor)
                    sellLevel = sellLevel - buffer
                End If
                ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, hkCandle, Trade.TradeExecutionDirection.Sell)
            End If
            'If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
            '    Dim buyLevel As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
            '    If buyLevel = Math.Round(hkCandle.High, 2) Then
            '        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(hkCandle.High, RoundOfType.Floor)
            '        buyLevel = buyLevel + buffer
            '    End If
            '    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, buyLevel, hkCandle, Trade.TradeExecutionDirection.Buy)
            'ElseIf hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
            '    Dim sellLevel As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
            '    If sellLevel = Math.Round(hkCandle.Low, 2) Then
            '        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(hkCandle.Low, RoundOfType.Floor)
            '        sellLevel = sellLevel - buffer
            '    End If
            '    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, sellLevel, hkCandle, Trade.TradeExecutionDirection.Sell)
            'End If
        End If
        Return ret
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _previousDayHighestATR <> Decimal.MinValue AndAlso _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)
            ret = (_previousDayHighestATR + todayHighestATR) / 2
            'ret = Math.Max(_previousDayHighestATR, todayHighestATR)
        End If
        Return ret
    End Function
End Class
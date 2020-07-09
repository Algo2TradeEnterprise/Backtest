Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HigherTimeframeDirectionMartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public HigherTimeframe As Integer
        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _xMinPayload As Dictionary(Of Date, Payload) = Nothing
    Private _emaPayload As Dictionary(Of Date, Decimal) = Nothing

    Private _slPoint As Decimal = Decimal.MinValue

    Private _previousDayHighestATR As Decimal = Decimal.MinValue

    Private ReadOnly _userInputs As StrategyRuleEntities

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)

        Dim firstCandleOfTheDay As Payload = Nothing
        For Each runningPayload In _signalPayload
            If runningPayload.Key.Date = _tradingDate.Date Then
                If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                    firstCandleOfTheDay = runningPayload.Value
                    Exit For
                End If
            End If
        Next
        If firstCandleOfTheDay IsNot Nothing AndAlso firstCandleOfTheDay.PreviousCandlePayload IsNot Nothing Then
            _previousDayHighestATR = _atrPayload.Max(Function(x)
                                                         If x.Key.Date >= firstCandleOfTheDay.PreviousCandlePayload.PayloadDate.Date Then
                                                             Return x.Value
                                                         Else
                                                             Return Decimal.MinValue
                                                         End If
                                                     End Function)
        End If
        If _previousDayHighestATR = Decimal.MinValue Then Throw New ApplicationException("Unable to calculate previous day higest atr")

        Dim exchangeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.ExchangeStartTime.Hours, _parentStrategy.ExchangeStartTime.Minutes, _parentStrategy.ExchangeStartTime.Seconds)
        _xMinPayload = Common.ConvertPayloadsToXMinutes(_inputPayload, _userInputs.HigherTimeframe, exchangeStartTime)
        Indicator.EMA.CalculateEMA(13, Payload.PayloadFields.Close, _xMinPayload, _emaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Me.EligibleToTakeTrade AndAlso Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) Then

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3
                    _slPoint = ConvertFloorCeling(GetAverageHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                ElseIf lastExecutedOrder IsNot Nothing Then
                    If signal.Item3.PayloadDate <> lastExecutedOrder.SignalCandle.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim slPoint As Decimal = _slPoint
                Dim iteration As Decimal = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                Dim multiplier As Integer = Math.Pow(2, iteration - 1)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, Math.Abs(_userInputs.MaxLossPerTrade) * -1 * multiplier, Trade.TypeOfStock.Cash)
                Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * multiplier * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - entryPrice

                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = entryPrice - slPoint,
                                .Target = entryPrice + targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = iteration
                            }
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = entryPrice + slPoint,
                                .Target = entryPrice - targetPoint,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                .Supporting2 = iteration
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
            Dim currentDirection As Trade.TradeExecutionDirection = GetHigherTimeframeDirection(currentTick)
            If currentDirection = currentTrade.EntryDirection Then
                Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
                Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
            Else
                ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
            End If
        End If
        Return ret
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

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _previousDayHighestATR <> Decimal.MinValue Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date >= _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)

            ret = (_previousDayHighestATR + todayHighestATR) / 2
        End If
        Return ret
    End Function

    Private Function GetHigherTimeframeDirection(ByVal currentTick As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        If currentTick IsNot Nothing Then
            Dim currentXMin As Date = _parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _xMinPayload, _userInputs.HigherTimeframe)
            If _xMinPayload.ContainsKey(currentXMin) Then
                Dim htCandle As Payload = _xMinPayload(currentXMin)
                If htCandle IsNot Nothing AndAlso htCandle.PreviousCandlePayload IsNot Nothing AndAlso htCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                    Dim ema As Decimal = _emaPayload(htCandle.PreviousCandlePayload.PayloadDate)
                    If htCandle.PreviousCandlePayload.Close < ema Then
                        ret = Trade.TradeExecutionDirection.Sell
                    Else
                        ret = Trade.TradeExecutionDirection.Buy
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim currentDirection As Trade.TradeExecutionDirection = GetHigherTimeframeDirection(currentTick)
            If currentDirection <> Trade.TradeExecutionDirection.None Then
                Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
                If currentDirection = Trade.TradeExecutionDirection.Buy AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                    If entryPrice = Math.Round(hkCandle.High, 2) Then
                        entryPrice = entryPrice + _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Buy)
                ElseIf currentDirection = Trade.TradeExecutionDirection.Sell AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                    Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                    If entryPrice = Math.Round(hkCandle.Low, 2) Then
                        entryPrice = entryPrice - _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

End Class
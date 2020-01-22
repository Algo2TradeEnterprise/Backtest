Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class DoubleTopDoubleBottomStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxStoplossPerStock As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private _fractalHighPayload As Dictionary(Of Date, Decimal)
    Private _fractalLowPayload As Dictionary(Of Date, Decimal)
    Private _swingHighPayload As Dictionary(Of Date, Decimal)
    Private _swingLowPayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)

    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _direction As Integer

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
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingHighPayload, _swingLowPayload)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
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
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, _parentStrategy.TradeType) AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            Dim signalCandle As Payload = Nothing
            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If lastExecutedTrade Is Nothing Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                ElseIf lastExecutedTrade.SignalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item5 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = 0
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item2
                    Dim slPrice As Decimal = signalCandleSatisfied.Item3
                    Dim targetPrice As Decimal = signalCandleSatisfied.Item4
                    Dim slPoint As Decimal = entryPrice - slPrice
                    Dim targetPoint As Decimal = targetPrice - entryPrice
                    If slPoint * 3 <= targetPoint Then
                        Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, slPrice, Math.Abs(_userInputs.MaxStoplossPerStock) * -1, _parentStrategy.StockType)

                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                    End If
                ElseIf signalCandleSatisfied.Item5 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = 0
                    Dim entryPrice As Decimal = signalCandleSatisfied.Item2
                    Dim slPrice As Decimal = signalCandleSatisfied.Item3
                    Dim targetPrice As Decimal = signalCandleSatisfied.Item4
                    Dim slPoint As Decimal = slPrice - entryPrice
                    Dim targetPoint As Decimal = entryPrice - targetPrice
                    If slPoint * 3 <= targetPoint Then
                        Dim quantity As Decimal = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, slPrice, entryPrice, Math.Abs(_userInputs.MaxStoplossPerStock) * -1, _parentStrategy.StockType)

                        parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = slPrice,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.Market,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
                    End If
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

    Private _lastSignalTime As Date = Date.MinValue
    Private _lastPreviousSignalTime As Date = Date.MinValue
    Private _lastSignalDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Private Function GetSignalCandle(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing Then
            If _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) > _fractalHighPayload(currentCandle.PayloadDate) Then
                Dim currentFractalU As Tuple(Of Date, Date) = GetFractalUFormingCandle(_fractalHighPayload, currentCandle.PayloadDate, 1)
                If currentFractalU IsNot Nothing AndAlso currentFractalU.Item2 = currentCandle.PayloadDate Then
                    Dim previousFractalU As Tuple(Of Date, Date) = GetFractalUFormingCandle(_fractalHighPayload, currentFractalU.Item1, 1)
                    If previousFractalU IsNot Nothing AndAlso previousFractalU.Item1.Date = _tradingDate.Date Then
                        Dim atr As Decimal = Math.Round(_atrPayload(currentCandle.PayloadDate), 4)
                        If Math.Abs(_signalPayload(currentFractalU.Item1).PreviousCandlePayload.High - _signalPayload(previousFractalU.Item1).PreviousCandlePayload.High) <= atr Then
                            _lastSignalTime = currentCandle.PayloadDate
                            _lastPreviousSignalTime = previousFractalU.Item2
                            _lastSignalDirection = Trade.TradeExecutionDirection.Sell
                        End If
                    End If
                End If
            End If
            _cts.Token.ThrowIfCancellationRequested()
            If _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) < _fractalLowPayload(currentCandle.PayloadDate) Then
                Dim currentFractalU As Tuple(Of Date, Date) = GetFractalUFormingCandle(_fractalLowPayload, currentCandle.PayloadDate, -1)
                If currentFractalU IsNot Nothing AndAlso currentFractalU.Item2 = currentCandle.PayloadDate Then
                    Dim previousFractalU As Tuple(Of Date, Date) = GetFractalUFormingCandle(_fractalLowPayload, currentFractalU.Item1, -1)
                    If previousFractalU IsNot Nothing AndAlso previousFractalU.Item1.Date = _tradingDate.Date Then
                        Dim atr As Decimal = Math.Round(_atrPayload(currentCandle.PayloadDate), 4)
                        If Math.Abs(_signalPayload(currentFractalU.Item1).PreviousCandlePayload.Low - _signalPayload(previousFractalU.Item1).PreviousCandlePayload.Low) <= atr Then
                            _lastSignalTime = currentCandle.PayloadDate
                            _lastPreviousSignalTime = previousFractalU.Item2
                            _lastSignalDirection = Trade.TradeExecutionDirection.Buy
                        End If
                    End If
                End If
            End If
            If _lastSignalTime <> Date.MinValue AndAlso _lastPreviousSignalTime <> Date.MinValue AndAlso _lastSignalDirection <> Trade.TradeExecutionDirection.None Then
                If _lastSignalDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = Math.Max(_swingLowPayload(currentCandle.PayloadDate), _fractalLowPayload(currentCandle.PayloadDate))
                    If currentCandle.Low > entryPrice AndAlso currentTick.Open <= entryPrice Then
                        Dim targetPrice As Decimal = entryPrice + ConvertFloorCeling(_atrPayload(currentCandle.PayloadDate) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                        Dim slPrice As Decimal = Math.Max(_fractalLowPayload(_signalPayload(_lastSignalTime).PreviousCandlePayload.PayloadDate), _fractalLowPayload(_signalPayload(_lastPreviousSignalTime).PreviousCandlePayload.PayloadDate))

                        ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection)(True, entryPrice, slPrice, targetPrice, Trade.TradeExecutionDirection.Buy)
                        _lastSignalTime = Date.MinValue
                        _lastPreviousSignalTime = Date.MinValue
                        _lastSignalDirection = Trade.TradeExecutionDirection.None
                    End If
                ElseIf _lastSignalDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = Math.Min(_swingHighPayload(currentCandle.PayloadDate), _fractalHighPayload(currentCandle.PayloadDate))
                    If currentCandle.Low < entryPrice AndAlso currentTick.Open >= entryPrice Then
                        Dim targetPrice As Decimal = entryPrice - ConvertFloorCeling(_atrPayload(currentCandle.PayloadDate) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Floor)
                        Dim slPrice As Decimal = Math.Min(_fractalHighPayload(_signalPayload(_lastSignalTime).PreviousCandlePayload.PayloadDate), _fractalHighPayload(_signalPayload(_lastPreviousSignalTime).PreviousCandlePayload.PayloadDate))

                        ret = New Tuple(Of Boolean, Decimal, Decimal, Decimal, Trade.TradeExecutionDirection)(True, entryPrice, slPrice, targetPrice, Trade.TradeExecutionDirection.Sell)
                        _lastSignalTime = Date.MinValue
                        _lastPreviousSignalTime = Date.MinValue
                        _lastSignalDirection = Trade.TradeExecutionDirection.None
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetFractalUFormingCandle(ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal beforeThisTime As Date, ByVal direction As Integer) As Tuple(Of Date, Date)
        Dim ret As Tuple(Of Date, Date) = Nothing
        If fractalPayload IsNot Nothing AndAlso fractalPayload.Count > 0 Then
            Dim checkingPayload As IEnumerable(Of KeyValuePair(Of Date, Decimal)) = fractalPayload.Where(Function(x)
                                                                                                             Return x.Key <= beforeThisTime
                                                                                                         End Function)
            If checkingPayload IsNot Nothing AndAlso checkingPayload.Count > 0 Then
                Dim firstCandleTime As Date = Date.MinValue
                Dim middleCandleTime As Date = Date.MinValue
                Dim lastCandleTime As Date = Date.MinValue
                For Each runningPayload In checkingPayload.OrderByDescending(Function(x)
                                                                                 Return x.Key
                                                                             End Function)
                    _cts.Token.ThrowIfCancellationRequested()
                    If direction > 0 Then
                        If firstCandleTime = Date.MinValue Then
                            firstCandleTime = runningPayload.Key
                        Else
                            If middleCandleTime = Date.MinValue Then
                                If fractalPayload(firstCandleTime) >= runningPayload.Value Then
                                    firstCandleTime = runningPayload.Key
                                Else
                                    middleCandleTime = runningPayload.Key
                                End If
                            Else
                                If fractalPayload(middleCandleTime) = runningPayload.Value Then
                                    middleCandleTime = runningPayload.Key
                                ElseIf fractalPayload(middleCandleTime) < runningPayload.Value Then
                                    firstCandleTime = middleCandleTime
                                    middleCandleTime = runningPayload.Key
                                ElseIf fractalPayload(middleCandleTime) > runningPayload.Value Then
                                    lastCandleTime = runningPayload.Key
                                    ret = New Tuple(Of Date, Date)(lastCandleTime, firstCandleTime)
                                    Exit For
                                End If
                            End If
                        End If
                    ElseIf direction < 0 Then
                        If firstCandleTime = Date.MinValue Then
                            firstCandleTime = runningPayload.Key
                        Else
                            If middleCandleTime = Date.MinValue Then
                                If fractalPayload(firstCandleTime) <= runningPayload.Value Then
                                    firstCandleTime = runningPayload.Key
                                Else
                                    middleCandleTime = runningPayload.Key
                                End If
                            Else
                                If fractalPayload(middleCandleTime) = runningPayload.Value Then
                                    middleCandleTime = runningPayload.Key
                                ElseIf fractalPayload(middleCandleTime) > runningPayload.Value Then
                                    firstCandleTime = middleCandleTime
                                    middleCandleTime = runningPayload.Key
                                ElseIf fractalPayload(middleCandleTime) < runningPayload.Value Then
                                    lastCandleTime = runningPayload.Key
                                    ret = New Tuple(Of Date, Date)(lastCandleTime, firstCandleTime)
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FravourableFractalPreviousLevelBreakoutAdaptiveMartingaleStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxProfitPerTrade As Decimal
        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _fractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _firstTime As Boolean

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
        _firstTime = True
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _fractalHighPayload, _fractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastOrder IsNot Nothing Then
                    Dim fractalCandleTime As Date = Date.ParseExact(lastOrder.Supporting4, "dd-MM-yyyy HH:mm:ss", Nothing)
                    If fractalCandleTime <> signal.Item6 Then
                        signalCandle = signal.Item4
                    End If
                Else
                    signalCandle = signal.Item4
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim stoploss As Decimal = signal.Item3

                Dim iteration As Integer = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                Dim lossPL As Decimal = _userInputs.MaxLossPerTrade + _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - Math.Abs(entryPrice - stoploss), lossPL, Trade.TypeOfStock.Cash)
                Dim profitPL As Decimal = _userInputs.MaxProfitPerTrade - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, signal.Item5, Trade.TypeOfStock.Cash)

                Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = signal.Item5,
                                                            .Quantity = quantity,
                                                            .Stoploss = stoploss,
                                                            .Target = target,
                                                            .Buffer = buffer,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.SL,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                            .Supporting2 = iteration,
                                                            .Supporting3 = Math.Round(_atrPayload(signalCandle.PayloadDate), 2),
                                                            .Supporting4 = signal.Item6.ToString("dd-MM-yyyy HH:mm:ss")
                                                        }

                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) = GetEntrySignal(currentCandle, currentTick)
            If signal IsNot Nothing Then
                If signal.Item1 Then
                    If currentTrade.EntryPrice <> signal.Item2 OrElse
                        currentTrade.PotentialStopLoss <> signal.Item3 OrElse
                        currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate OrElse
                        currentTrade.EntryDirection <> signal.Item5 Then
                        ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                    End If
                End If
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

    Private _lastSignalHighLevel As Payload
    Private _lastSignalLowLevel As Payload
    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(currentCandle.PreviousCandlePayload.Close, RoundOfType.Floor)
            Dim buyLevel As Decimal = Decimal.MinValue
            Dim sellLevel As Decimal = Decimal.MinValue

            Dim atr As Decimal = ConvertFloorCeling(_atrPayload(currentCandle.PreviousCandlePayload.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
            If _firstTime Then
                _firstTime = False
                Dim preHighFractal As Payload = GetPreviousFractalCandle(_fractalHighPayload, currentCandle.PreviousCandlePayload)
                If preHighFractal IsNot Nothing Then
                    If _fractalHighPayload(preHighFractal.PreviousCandlePayload.PayloadDate) > _fractalHighPayload(preHighFractal.PayloadDate) AndAlso
                        _fractalHighPayload(preHighFractal.PreviousCandlePayload.PayloadDate) - _fractalHighPayload(preHighFractal.PayloadDate) <= atr Then
                        _lastSignalHighLevel = preHighFractal
                        buyLevel = _fractalHighPayload(_lastSignalHighLevel.PreviousCandlePayload.PayloadDate) + buffer
                    End If
                End If
                Dim preLowFractal As Payload = GetPreviousFractalCandle(_fractalLowPayload, currentCandle.PreviousCandlePayload)
                If preLowFractal IsNot Nothing Then
                    If _fractalLowPayload(preLowFractal.PreviousCandlePayload.PayloadDate) < _fractalLowPayload(preLowFractal.PayloadDate) AndAlso
                        _fractalLowPayload(preLowFractal.PayloadDate) - _fractalLowPayload(preLowFractal.PreviousCandlePayload.PayloadDate) <= atr Then
                        _lastSignalLowLevel = preLowFractal
                        sellLevel = _fractalLowPayload(_lastSignalLowLevel.PreviousCandlePayload.PayloadDate) + buffer
                    End If
                End If
            Else
                If _fractalHighPayload(currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate) > _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    _lastSignalHighLevel = Nothing
                End If
                If _lastSignalHighLevel Is Nothing Then
                    If _fractalHighPayload(currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate) > _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                        _fractalHighPayload(currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate) - _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) <= atr Then
                        _lastSignalHighLevel = currentCandle.PreviousCandlePayload
                        buyLevel = _fractalHighPayload(_lastSignalHighLevel.PreviousCandlePayload.PayloadDate) + buffer
                    End If
                Else
                    buyLevel = _fractalHighPayload(_lastSignalHighLevel.PreviousCandlePayload.PayloadDate) + buffer
                End If

                If _fractalLowPayload(currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate) < _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    _lastSignalLowLevel = Nothing
                End If
                If _lastSignalLowLevel Is Nothing Then
                    If _fractalLowPayload(currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate) < _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) AndAlso
                        _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) - _fractalLowPayload(currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate) <= atr Then
                        _lastSignalLowLevel = currentCandle.PreviousCandlePayload
                        sellLevel = _fractalLowPayload(_lastSignalLowLevel.PreviousCandlePayload.PayloadDate) - buffer
                    End If
                Else
                    sellLevel = _fractalLowPayload(_lastSignalLowLevel.PreviousCandlePayload.PayloadDate) - buffer
                End If
            End If

            If buyLevel <> Decimal.MinValue AndAlso sellLevel <> Decimal.MinValue Then
                Dim avg As Decimal = (buyLevel + sellLevel) / 2
                Dim range As Decimal = buyLevel - avg
                If currentTick.Open >= avg + range * 60 / 100 Then
                    If currentCandle.PreviousCandlePayload.Open < _fractalHighPayload(_lastSignalHighLevel.PayloadDate) OrElse
                        currentCandle.PreviousCandlePayload.Close < _fractalHighPayload(_lastSignalHighLevel.PayloadDate) Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) _
                        (True, buyLevel, currentCandle.PreviousCandlePayload.Low - buffer, currentCandle.PreviousCandlePayload,
                         Trade.TradeExecutionDirection.Buy, _lastSignalHighLevel.PayloadDate)
                    End If
                ElseIf currentTick.Open <= avg - range * 60 / 100 Then
                    If currentCandle.PreviousCandlePayload.Open > _fractalLowPayload(_lastSignalLowLevel.PayloadDate) OrElse
                        currentCandle.PreviousCandlePayload.Close > _fractalLowPayload(_lastSignalLowLevel.PayloadDate) Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) _
                        (True, sellLevel, currentCandle.PreviousCandlePayload.High + buffer, currentCandle.PreviousCandlePayload,
                         Trade.TradeExecutionDirection.Sell, _lastSignalLowLevel.PayloadDate)
                    End If
                End If
            ElseIf buyLevel <> Decimal.MinValue Then
                If currentCandle.PreviousCandlePayload.Open < _fractalHighPayload(_lastSignalHighLevel.PayloadDate) OrElse
                        currentCandle.PreviousCandlePayload.Close < _fractalHighPayload(_lastSignalHighLevel.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) _
                        (True, buyLevel, currentCandle.PreviousCandlePayload.Low - buffer, currentCandle.PreviousCandlePayload,
                         Trade.TradeExecutionDirection.Buy, _lastSignalHighLevel.PayloadDate)
                End If
            ElseIf sellLevel <> Decimal.MinValue Then
                If currentCandle.PreviousCandlePayload.Open > _fractalLowPayload(_lastSignalLowLevel.PayloadDate) OrElse
                        currentCandle.PreviousCandlePayload.Close > _fractalLowPayload(_lastSignalLowLevel.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) _
                        (True, sellLevel, currentCandle.PreviousCandlePayload.High + buffer, currentCandle.PreviousCandlePayload,
                         Trade.TradeExecutionDirection.Sell, _lastSignalLowLevel.PayloadDate)
                End If
            End If

            If ret IsNot Nothing AndAlso ret.Item1 Then
                Dim tgtSl As Tuple(Of Boolean, Date) = IsTargetStoplossHit(ret.Item2, Math.Abs(ret.Item2 - ret.Item3), ret.Item6, currentCandle.PayloadDate, ret.Item5)
                If tgtSl IsNot Nothing Then
                    If tgtSl.Item1 Then
                        ret = Nothing
                    Else
                        If tgtSl.Item2 <> Date.MinValue Then
                            Dim entry As Decimal = ret.Item2
                            Dim stoploss As Decimal = ret.Item3
                            Dim direction As Trade.TradeExecutionDirection = ret.Item5
                            Dim fractalTime As Date = ret.Item6
                            Dim signalCandle As Payload = Nothing
                            For Each runningPayload In _signalPayload.OrderByDescending(Function(x)
                                                                                            Return x.Key
                                                                                        End Function)
                                If runningPayload.Key > fractalTime AndAlso runningPayload.Key < tgtSl.Item2 Then
                                    If direction = Trade.TradeExecutionDirection.Buy Then
                                        If runningPayload.Value.Open < _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) OrElse
                                            runningPayload.Value.Close < _fractalHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                                            signalCandle = runningPayload.Value
                                            stoploss = signalCandle.Low - buffer
                                            Exit For
                                        End If
                                    ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                                        If runningPayload.Value.Open > _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) OrElse
                                            runningPayload.Value.Close > _fractalLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                                            signalCandle = runningPayload.Value
                                            stoploss = signalCandle.High + buffer
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                            If signalCandle IsNot Nothing Then
                                ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, Date) _
                                      (True, entry, stoploss, signalCandle, direction, fractalTime)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetPreviousFractalCandle(ByVal fractalPayload As Dictionary(Of Date, Decimal), ByVal candle As Payload) As Payload
        Dim ret As Payload = Nothing
        Dim currentFractal As Decimal = fractalPayload(candle.PayloadDate)
        For Each runningPayload In fractalPayload.OrderByDescending(Function(x)
                                                                        Return x.Key
                                                                    End Function)
            If runningPayload.Key <= candle.PayloadDate Then
                If fractalPayload(_signalPayload(runningPayload.Key).PreviousCandlePayload.PayloadDate) <> currentFractal Then
                    ret = _signalPayload(runningPayload.Key)
                    Exit For
                End If
            End If
        Next
        Return ret
    End Function

    Private Function IsTargetStoplossHit(ByVal entryPrice As Decimal, ByVal slPoint As Decimal,
                                         ByVal startTime As Date, ByVal endTime As Date,
                                         ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Date)
        Dim ret As Tuple(Of Boolean, Date) = Nothing
        Dim signalTriggerd As Boolean = False
        Dim signalTriggerTime As Date = Date.MinValue
        For Each runningPayload In _signalPayload
            If runningPayload.Key > startTime AndAlso runningPayload.Key < endTime Then
                If Not signalTriggerd Then
                    If direction = Trade.TradeExecutionDirection.Buy AndAlso
                        runningPayload.Value.High >= entryPrice Then
                        signalTriggerd = True
                        signalTriggerTime = runningPayload.Key
                        ret = New Tuple(Of Boolean, Date)(False, signalTriggerTime)
                    ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso
                        runningPayload.Value.Low <= entryPrice Then
                        signalTriggerd = True
                        signalTriggerTime = runningPayload.Key
                        ret = New Tuple(Of Boolean, Date)(False, signalTriggerTime)
                    End If
                End If
                If signalTriggerd Then
                    If runningPayload.Value.High >= entryPrice + slPoint OrElse
                        runningPayload.Value.Low <= entryPrice - slPoint Then
                        ret = New Tuple(Of Boolean, Date)(True, signalTriggerTime)
                        Exit For
                    End If
                End If
            End If
        Next
        Return ret
    End Function
End Class
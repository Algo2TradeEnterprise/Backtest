Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowSLStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public BreakevenMultiplier As Decimal
        Public StoplossMakeupTrade As Boolean
        Public NumberOfTrade As Integer
        Public ModifyCandleTarget As Boolean
        Public ModifyNumberOfTrade As Boolean
        Public MaxLossPercentageOfCapital As Decimal
    End Class
#End Region

    Private _potentialHighEntryPrice As Decimal = Decimal.MinValue
    Private _potentialLowEntryPrice As Decimal = Decimal.MinValue
    Private _signalCandle As Payload
    Private _entryChanged As Boolean = False
    Private _ATRPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _FractalHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _FractalLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _userInputs As StrategyRuleEntities
    Private _entryRemark As String = ""
    Private _targetPoint As Decimal = Decimal.MinValue
    Private ReadOnly _stockATR As Decimal
    Private ReadOnly _dayATR As Decimal
    Private ReadOnly _slPoint As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal stockATR As Decimal,
                   ByVal dayATR As Decimal,
                   ByVal slPoint As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _stockATR = stockATR
        _dayATR = dayATR
        _slPoint = slPoint
        _userInputs = New StrategyRuleEntities With {
            .TargetMultiplier = CType(_entities, StrategyRuleEntities).TargetMultiplier,
            .BreakevenMovement = CType(_entities, StrategyRuleEntities).BreakevenMovement,
            .BreakevenMultiplier = CType(_entities, StrategyRuleEntities).BreakevenMultiplier,
            .StoplossMakeupTrade = CType(_entities, StrategyRuleEntities).StoplossMakeupTrade,
            .ModifyCandleTarget = CType(_entities, StrategyRuleEntities).ModifyCandleTarget,
            .ModifyNumberOfTrade = CType(_entities, StrategyRuleEntities).ModifyNumberOfTrade,
            .MaxLossPercentageOfCapital = CType(_entities, StrategyRuleEntities).MaxLossPercentageOfCapital,
            .NumberOfTrade = _parentStrategy.NumberOfTradesPerStockPerDay
        }
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _ATRPayload)
        Indicator.FractalBands.CalculateFractal(_signalPayload, _FractalHighPayload, _FractalLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)

        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < _userInputs.NumberOfTrade AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing

            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedTrade Is Nothing Then
                    signalCandle = _signalCandle
                Else
                    If lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso lastExecutedTrade.PLPoint > 0 Then
                        'Breakeven exit
                        If lastExecutedTrade.EntryDirection = signalCandleSatisfied.Item4 Then
                            If IsSignalTriggered(signalCandleSatisfied.Item3, If(signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy, Trade.TradeExecutionDirection.Sell, Trade.TradeExecutionDirection.Buy), lastExecutedTrade.ExitTime, currentTick.PayloadDate) Then
                                signalCandle = _signalCandle
                            End If
                        Else
                            signalCandle = _signalCandle
                        End If
                    Else
                        signalCandle = _signalCandle
                    End If
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandleSatisfied.Item2 + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandleSatisfied.Item3,
                        .Target = .EntryPrice + ConvertFloorCeling(_targetPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = _entryRemark,
                        .Supporting3 = _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) + 1,
                        .Supporting4 = _userInputs.TargetMultiplier,
                        .Supporting5 = _userInputs.NumberOfTrade
                    }
                ElseIf signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandleSatisfied.Item2 - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandleSatisfied.Item3,
                        .Target = .EntryPrice - ConvertFloorCeling(_targetPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = _entryRemark,
                        .Supporting3 = _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) + 1,
                        .Supporting4 = _userInputs.TargetMultiplier,
                        .Supporting5 = _userInputs.NumberOfTrade
                    }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            'Stoploss makeup target calculation
            If _userInputs.StoplossMakeupTrade AndAlso _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) >= _userInputs.NumberOfTrade - 1 Then
                Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
                If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then
                    Dim totalLoss As Decimal = closeTrades.Sum(Function(x)
                                                                   Return x.PLAfterBrokerage
                                                               End Function)
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(currentTick.TradingSymbol, parameter.EntryPrice, parameter.Quantity, Math.Abs(totalLoss), parameter.EntryDirection, _parentStrategy.StockType)
                    If targetPrice <> Decimal.MinValue Then parameter.Target = targetPrice
                End If
            End If

            'Quantity calculation
            If parameter.EntryPrice * parameter.Quantity / _parentStrategy.MarginMultiplier < 15000 Then
                parameter.Quantity = 2 * _lotSize
            End If

            'Stop taking trade if loss is greater that x% of capital
            If _userInputs.MaxLossPercentageOfCapital <> Decimal.MinValue Then
                Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
                If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then
                    Dim totalLoss As Decimal = closeTrades.Sum(Function(x)
                                                                   Return x.PLAfterBrokerage
                                                               End Function)

                    If Math.Abs(totalLoss) >= (parameter.EntryPrice * parameter.Quantity / _parentStrategy.MarginMultiplier) * _userInputs.MaxLossPercentageOfCapital / 100 AndAlso
                        _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < 0 Then
                        'Neglect this trade
                        'Console.WriteLine(String.Format("Trade neglected. Time: {0}, Direction:{1}, Symbol:{2}", currentTick.PayloadDate, parameter.EntryDirection, _tradingSymbol))
                        Me.EligibleToTakeTrade = False
                    Else
                        ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                    End If
                Else
                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If
            Else
                ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
            End If

            'ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})

            If _parentStrategy.StockMaxProfitPercentagePerDay <> Decimal.MaxValue AndAlso Me.MaxProfitOfThisStock = Decimal.MaxValue Then
                Me.MaxProfitOfThisStock = _parentStrategy.CalculatePL(currentTick.TradingSymbol, parameter.EntryPrice, ConvertFloorCeling(parameter.EntryPrice + parameter.EntryPrice * _parentStrategy.StockMaxProfitPercentagePerDay / 100, _parentStrategy.TickSize, RoundOfType.Celing), parameter.Quantity, _lotSize, _parentStrategy.StockType)
            End If
            If _parentStrategy.StockMaxLossPercentagePerDay <> Decimal.MinValue AndAlso Me.MaxLossOfThisStock = Decimal.MinValue Then
                Me.MaxLossOfThisStock = _parentStrategy.CalculatePL(currentTick.TradingSymbol, parameter.EntryPrice, ConvertFloorCeling(parameter.EntryPrice - parameter.EntryPrice * _parentStrategy.StockMaxLossPercentagePerDay / 100, _parentStrategy.TickSize, RoundOfType.Celing), parameter.Quantity, _lotSize, _parentStrategy.StockType)
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    Dim entryPrice As Decimal = Decimal.MinValue
                    If signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy Then
                        entryPrice = signalCandleSatisfied.Item2 + buffer
                    ElseIf signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Sell Then
                        entryPrice = signalCandleSatisfied.Item2 - buffer
                    End If
                    If (currentTrade.EntryDirection <> signalCandleSatisfied.Item4) OrElse
                        (currentTrade.EntryPrice <> entryPrice) Then
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
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim excpectedTarget As Decimal = currentTrade.EntryPrice + (currentTrade.PotentialTarget - currentTrade.EntryPrice) * _userInputs.BreakevenMultiplier
                If currentTick.Open >= excpectedTarget Then
                    triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, _lotSize, _parentStrategy.StockType)
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim excpectedTarget As Decimal = currentTrade.EntryPrice - (currentTrade.EntryPrice - currentTrade.PotentialTarget) * _userInputs.BreakevenMultiplier
                If currentTick.Open <= excpectedTarget Then
                    triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, _lotSize, _parentStrategy.StockType)
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move to breakeven: {0}. Time:{1}", triggerPrice, currentTick.PayloadDate))
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        If Not _entryChanged AndAlso _slPoint <> Decimal.MinValue AndAlso
            _potentialHighEntryPrice <> Decimal.MinValue AndAlso _potentialLowEntryPrice <> Decimal.MinValue Then
            Dim highBuffer As Decimal = _parentStrategy.CalculateBuffer(_potentialHighEntryPrice, RoundOfType.Floor)
            Dim lowBuffer As Decimal = _parentStrategy.CalculateBuffer(_potentialLowEntryPrice, RoundOfType.Floor)
            If currentTick.High >= _potentialHighEntryPrice + highBuffer Then
                _potentialHighEntryPrice = _potentialHighEntryPrice + highBuffer
                _potentialLowEntryPrice = _potentialHighEntryPrice - 2 * _slPoint
                _entryChanged = True
            ElseIf currentTick.Low <= _potentialLowEntryPrice - lowBuffer Then
                _potentialLowEntryPrice = _potentialLowEntryPrice - lowBuffer
                _potentialHighEntryPrice = _potentialLowEntryPrice + 2 * _slPoint
                _entryChanged = True
            End If
        End If
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso
            Not candle.DeadCandle AndAlso Not candle.PreviousCandlePayload.DeadCandle Then
            If _potentialHighEntryPrice = Decimal.MinValue AndAlso _potentialLowEntryPrice = Decimal.MinValue Then
                If IsSignalCandle(candle) Then
                    _potentialHighEntryPrice = candle.High
                    _potentialLowEntryPrice = candle.Low
                    _signalCandle = candle
                    Dim atr As Decimal = ConvertFloorCeling(_ATRPayload(_signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
                    If atr * _userInputs.TargetMultiplier >= _slPoint * _userInputs.TargetMultiplier Then
                        _targetPoint = atr * _userInputs.TargetMultiplier
                        If _targetPoint > _dayATR / 2 Then
                            _potentialHighEntryPrice = Decimal.MinValue
                            _potentialLowEntryPrice = Decimal.MinValue
                            _signalCandle = Nothing
                        End If
                    Else
                        _targetPoint = _slPoint * _userInputs.TargetMultiplier
                    End If
                End If
            End If

            If _potentialHighEntryPrice <> Decimal.MinValue AndAlso _potentialLowEntryPrice <> Decimal.MinValue Then
                If _entryChanged Then
                    Dim middlePoint As Decimal = (_potentialHighEntryPrice + _potentialLowEntryPrice) / 2
                    Dim range As Decimal = _potentialHighEntryPrice - middlePoint
                    If currentTick.Open >= middlePoint + range * 60 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialHighEntryPrice, middlePoint, Trade.TradeExecutionDirection.Buy)
                    ElseIf currentTick.Open <= middlePoint - range * 60 / 100 Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialLowEntryPrice, middlePoint, Trade.TradeExecutionDirection.Sell)
                    End If
                Else
                    Dim tradeDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                    Dim middlePoint As Decimal = (_potentialHighEntryPrice + _potentialLowEntryPrice) / 2
                    Dim range As Decimal = _potentialHighEntryPrice - middlePoint
                    Dim buffer As Decimal = Decimal.MinValue
                    If currentTick.Open >= middlePoint + range * 30 / 100 Then
                        tradeDirection = Trade.TradeExecutionDirection.Buy
                        buffer = _parentStrategy.CalculateBuffer(_potentialHighEntryPrice, RoundOfType.Floor)
                    ElseIf currentTick.Open <= middlePoint - range * 30 / 100 Then
                        tradeDirection = Trade.TradeExecutionDirection.Sell
                        buffer = _parentStrategy.CalculateBuffer(_potentialLowEntryPrice, RoundOfType.Floor)
                    End If
                    If tradeDirection = Trade.TradeExecutionDirection.Buy Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialHighEntryPrice, _potentialHighEntryPrice + buffer - _slPoint, Trade.TradeExecutionDirection.Buy)
                    ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialLowEntryPrice, _potentialLowEntryPrice - buffer + _slPoint, Trade.TradeExecutionDirection.Sell)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        If IsFractalHighChanged(candle.PayloadDate) AndAlso candle.Low < candle.PreviousCandlePayload.Low Then
            ret = True
        ElseIf IsFractalLowChanged(candle.PayloadDate) AndAlso candle.High > candle.PreviousCandlePayload.High Then
            ret = True
        End If
        Return ret
    End Function

    Private Function IsFractalHighChanged(ByVal currentTime As Date) As Boolean
        Dim ret As Boolean = False
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key <= currentTime Then
                    If _FractalHighPayload(runningPayload.Value.PayloadDate) <> _FractalHighPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate) Then
                        ret = True
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Private Function IsFractalLowChanged(ByVal currentTime As Date) As Boolean
        Dim ret As Boolean = False
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date AndAlso runningPayload.Key <= currentTime Then
                    If _FractalLowPayload(runningPayload.Value.PayloadDate) <> _FractalLowPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate) Then
                        ret = True
                        Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function
End Class

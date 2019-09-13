Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FixedLevelBasedStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public StoplossMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public BreakevenMultiplier As Decimal
        Public StoplossMakeupTrade As Integer
        Public NumberOfTrade As Integer
        Public LevelType As TypeOfLevel

        Enum TypeOfLevel
            None = 1
            BreakoutATR
            StoplossATR
            Candle
        End Enum
    End Class
#End Region

    Private _potentialHighEntryPrice As Decimal = Decimal.MinValue
    Private _potentialLowEntryPrice As Decimal = Decimal.MinValue
    Private _signalCandle As Payload
    Private _entryChanged As Boolean = False
    Private _ATRPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _userInputs As StrategyRuleEntities
    Private _entryRemark As String = ""
    Private ReadOnly _stockATR As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal stockATR As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _stockATR = stockATR
        _userInputs = New StrategyRuleEntities With {
            .TargetMultiplier = CType(_entities, StrategyRuleEntities).TargetMultiplier,
            .StoplossMultiplier = CType(_entities, StrategyRuleEntities).StoplossMultiplier,
            .BreakevenMovement = CType(_entities, StrategyRuleEntities).BreakevenMovement,
            .BreakevenMultiplier = CType(_entities, StrategyRuleEntities).BreakevenMultiplier,
            .StoplossMakeupTrade = CType(_entities, StrategyRuleEntities).StoplossMakeupTrade,
            .LevelType = CType(_entities, StrategyRuleEntities).LevelType,
            .NumberOfTrade = _parentStrategy.NumberOfTradesPerStockPerDay
        }
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _ATRPayload)
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
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime Then
            Dim signalCandle As Payload = Nothing

            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.BreakoutATR OrElse
                    (_userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.StoplossATR AndAlso _entryChanged) OrElse
                    _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Candle Then
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
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandleSatisfied.Item2 + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandleSatisfied.Item3 - If(_userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Candle, buffer, 0),
                        .Target = .EntryPrice + ConvertFloorCeling((.EntryPrice - .Stoploss) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = _entryRemark,
                        .Supporting3 = _userInputs.TargetMultiplier
                    }
                ElseIf signalCandleSatisfied.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandleSatisfied.Item2 - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandleSatisfied.Item3 + If(_userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Candle, buffer, 0),
                        .Target = .EntryPrice - ConvertFloorCeling((.Stoploss - .EntryPrice) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = _entryRemark,
                        .Supporting3 = _userInputs.TargetMultiplier
                    }
                End If
            End If
        End If
        If parameter IsNot Nothing Then
            If _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) >= _userInputs.StoplossMakeupTrade - 1 Then
                Dim closeTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionStatus.Close)
                If closeTrades IsNot Nothing AndAlso closeTrades.Count > 0 Then
                    Dim totalLoss As Decimal = closeTrades.Sum(Function(x)
                                                                   Return x.PLAfterBrokerage
                                                               End Function)
                    Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(currentTick.TradingSymbol, parameter.EntryPrice, parameter.Quantity, Math.Abs(totalLoss), parameter.EntryDirection, _parentStrategy.StockType)
                    If targetPrice <> Decimal.MinValue Then parameter.Target = targetPrice
                End If
            End If
            Dim capital As Decimal = parameter.EntryPrice * parameter.Quantity / _parentStrategy.MarginMultiplier
            If capital < 15000 Then
                parameter.Quantity = 2 * _lotSize
            End If
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
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
        If Not _entryChanged AndAlso GetSignalCandleATR() <> Decimal.MinValue AndAlso
            _potentialHighEntryPrice <> Decimal.MinValue AndAlso _potentialLowEntryPrice <> Decimal.MinValue Then
            Dim highBuffer As Decimal = _parentStrategy.CalculateBuffer(_potentialHighEntryPrice, RoundOfType.Floor)
            Dim lowBuffer As Decimal = _parentStrategy.CalculateBuffer(_potentialLowEntryPrice, RoundOfType.Floor)
            If currentTick.High >= _potentialHighEntryPrice + highBuffer Then
                If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.StoplossATR Then
                    Dim middlePoint As Decimal = _potentialHighEntryPrice
                    _potentialHighEntryPrice = middlePoint + ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    _potentialLowEntryPrice = middlePoint - ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                ElseIf _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.BreakoutATR Then
                    _potentialLowEntryPrice = _potentialHighEntryPrice - 2 * ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                ElseIf _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Candle Then
                    _potentialLowEntryPrice = _potentialHighEntryPrice - 2 * ConvertFloorCeling(_signalCandle.CandleRange * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                End If
                _entryChanged = True
            ElseIf currentTick.Low <= _potentialLowEntryPrice - lowBuffer Then
                If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.StoplossATR Then
                    Dim middlePoint As Decimal = _potentialLowEntryPrice
                    _potentialHighEntryPrice = middlePoint + ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    _potentialLowEntryPrice = middlePoint - ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                ElseIf _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.BreakoutATR Then
                    _potentialHighEntryPrice = _potentialLowEntryPrice + 2 * ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                ElseIf _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Candle Then
                    _potentialHighEntryPrice = _potentialLowEntryPrice + 2 * ConvertFloorCeling(_signalCandle.CandleRange * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                End If
                _entryChanged = True
            End If
        End If
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso
            Not candle.DeadCandle AndAlso Not candle.PreviousCandlePayload.DeadCandle Then
            If _potentialHighEntryPrice = Decimal.MinValue AndAlso _potentialLowEntryPrice = Decimal.MinValue Then
                If IsCandleHalf(candle) AndAlso candle.CandleRange > _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor) Then
                    _potentialHighEntryPrice = candle.High
                    _potentialLowEntryPrice = candle.Low
                    _signalCandle = candle
                    If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.None Then
                        _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Candle
                    End If
                    _entryRemark = "Candle Half"
                ElseIf IsPinBar(candle) Then
                    _potentialHighEntryPrice = candle.High
                    _potentialLowEntryPrice = candle.Low
                    _signalCandle = candle
                    If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.None Then
                        _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.BreakoutATR
                        _userInputs.TargetMultiplier = Math.Floor(_userInputs.TargetMultiplier - _userInputs.TargetMultiplier * 25 / 100)
                    End If
                    _entryRemark = "Pin Bar"
                ElseIf IsTweezerPattern(candle) Then
                    _potentialHighEntryPrice = candle.High
                    _potentialLowEntryPrice = candle.Low
                    _signalCandle = candle
                    If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.None Then
                        _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.BreakoutATR
                        _userInputs.TargetMultiplier = Math.Floor(_userInputs.TargetMultiplier - _userInputs.TargetMultiplier * 25 / 100)
                    End If
                    _entryRemark = "Tweezer Pattern"
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
                    If currentTick.Open >= middlePoint + range * 30 / 100 Then
                        tradeDirection = Trade.TradeExecutionDirection.Buy
                    ElseIf currentTick.Open <= middlePoint - range * 30 / 100 Then
                        tradeDirection = Trade.TradeExecutionDirection.Sell
                    End If
                    Select Case _userInputs.LevelType
                        Case StrategyRuleEntities.TypeOfLevel.BreakoutATR
                            If tradeDirection = Trade.TradeExecutionDirection.Buy Then
                                ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialHighEntryPrice, _potentialHighEntryPrice - ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing), Trade.TradeExecutionDirection.Buy)
                            ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
                                ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialLowEntryPrice, _potentialLowEntryPrice + ConvertFloorCeling(GetSignalCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing), Trade.TradeExecutionDirection.Sell)
                            End If
                        Case StrategyRuleEntities.TypeOfLevel.Candle
                            If tradeDirection = Trade.TradeExecutionDirection.Buy Then
                                ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialHighEntryPrice, _potentialLowEntryPrice, Trade.TradeExecutionDirection.Buy)
                            ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
                                ret = New Tuple(Of Boolean, Decimal, Decimal, Trade.TradeExecutionDirection)(True, _potentialLowEntryPrice, _potentialHighEntryPrice, Trade.TradeExecutionDirection.Sell)
                            End If
                    End Select
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetSignalCandleATR() As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _signalCandle IsNot Nothing Then
            If _ATRPayload IsNot Nothing AndAlso _ATRPayload.Count > 0 AndAlso _ATRPayload.ContainsKey(_signalCandle.PayloadDate) Then
                ret = ConvertFloorCeling(_ATRPayload(_signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
            End If
        End If
        Return ret
    End Function

    Private Function IsTweezerPattern(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        If candle.PreviousCandlePayload.CandleColor = Color.Red Then
            If candle.PreviousCandlePayload.CandleWicks.Top <= candle.PreviousCandlePayload.CandleRange * 50 / 100 AndAlso
                candle.PreviousCandlePayload.CandleWicks.Bottom <= candle.PreviousCandlePayload.CandleRange * 25 / 100 Then
                If candle.CandleColor = Color.Green AndAlso candle.High >= candle.PreviousCandlePayload.High Then
                    If candle.CandleWicks.Top <= candle.CandleRange * 25 / 100 AndAlso
                        candle.CandleWicks.Bottom <= candle.CandleRange * 50 / 100 Then
                        ret = True
                    End If
                End If
            End If
        ElseIf candle.PreviousCandlePayload.CandleColor = Color.Green Then
            If candle.PreviousCandlePayload.CandleWicks.Top <= candle.PreviousCandlePayload.CandleRange * 25 / 100 AndAlso
                candle.PreviousCandlePayload.CandleWicks.Bottom <= candle.PreviousCandlePayload.CandleRange * 50 / 100 Then
                If candle.CandleColor = Color.Red AndAlso candle.Low <= candle.PreviousCandlePayload.Low Then
                    If candle.CandleWicks.Top <= candle.CandleRange * 50 / 100 AndAlso
                        candle.CandleWicks.Bottom <= candle.CandleRange * 25 / 100 Then
                        ret = True
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function IsCandleHalf(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, 9, 17, 0)
        Dim middlePoint As Decimal = (candle.PreviousCandlePayload.High + candle.PreviousCandlePayload.Low) / 2
        If candle.PayloadDate >= tradeStartTime AndAlso candle.CandleRange <= candle.PreviousCandlePayload.CandleRange / 2 AndAlso
            (candle.High <= ConvertFloorCeling(middlePoint, _parentStrategy.TickSize, RoundOfType.Floor) OrElse
            candle.Low >= ConvertFloorCeling(middlePoint, _parentStrategy.TickSize, RoundOfType.Celing)) Then
            ret = True
        End If
        Return ret
    End Function

    Private Function IsPinBar(ByVal candle As Payload) As Boolean
        Dim ret As Boolean = False
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(candle.PayloadDate.AddMinutes(_parentStrategy.SignalTimeFrame), _signalPayload))
        Dim candleHighBuffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
        Dim candleLowBuffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
        If candle.CandleWicks.Bottom >= ConvertFloorCeling(candle.CandleRange * 50 / 100, _parentStrategy.TickSize, RoundOfType.Celing) AndAlso
            candle.Volume >= candle.PreviousCandlePayload.Volume AndAlso
            candle.Low < candle.PreviousCandlePayload.Low AndAlso
            currentMinuteCandlePayload.High >= candle.High + candleHighBuffer Then
            If candle.High < candle.PreviousCandlePayload.High Then
                ret = True
            Else
                Dim dayLow As Decimal = _signalPayload.Values.Min(Function(x)
                                                                      If x.PayloadDate.Date = _tradingDate.Date AndAlso
                                                                        x.PayloadDate <= candle.PayloadDate Then
                                                                          Return x.Low
                                                                      Else
                                                                          Return Decimal.MaxValue
                                                                      End If
                                                                  End Function)
                If candle.Low = dayLow Then
                    ret = True
                End If
            End If
        ElseIf candle.CandleWicks.Top >= ConvertFloorCeling(candle.CandleRange * 50 / 100, _parentStrategy.TickSize, RoundOfType.Celing) AndAlso
            candle.Volume >= candle.PreviousCandlePayload.Volume AndAlso
            candle.High > candle.PreviousCandlePayload.High AndAlso
            currentMinuteCandlePayload.Low <= candle.Low - candleLowBuffer Then
            If candle.Low > candle.PreviousCandlePayload.Low Then
                ret = True
            Else
                Dim dayHigh As Decimal = _signalPayload.Values.Max(Function(x)
                                                                       If x.PayloadDate.Date = _tradingDate.Date AndAlso
                                                                        x.PayloadDate <= candle.PayloadDate Then
                                                                           Return x.High
                                                                       Else
                                                                           Return Decimal.MinValue
                                                                       End If
                                                                   End Function)
                If candle.High = dayHigh Then
                    ret = True
                End If
            End If
        End If
        Return ret
    End Function
End Class

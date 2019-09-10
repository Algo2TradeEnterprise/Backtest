Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class ATRFixedLevelBasedStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public StoplossMultiplier As Decimal
        Public BreakevenMovement As Boolean
        Public BreakevenMultiplier As Decimal
        Public LevelType As TypeOfLevel

        Enum TypeOfLevel
            None = 1
            Breakout
            SL
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
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > Math.Abs(_parentStrategy.OverAllLossPerDay) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso GetLastDayLastCandleATR() <> Decimal.MinValue Then
            Dim signalCandle As Payload = Nothing

            Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Breakout OrElse
                    (_userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.SL AndAlso _entryChanged) OrElse
                    _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Candle Then
                    signalCandle = _signalCandle
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim tradeDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                Dim middlePoint As Decimal = (signalCandleSatisfied.Item2 + signalCandleSatisfied.Item3) / 2
                Dim range As Decimal = signalCandleSatisfied.Item2 - middlePoint
                If currentTick.Open >= middlePoint + range * 60 / 100 Then
                    tradeDirection = Trade.TradeExecutionDirection.Buy
                ElseIf currentTick.Open <= middlePoint - range * 60 / 100 Then
                    tradeDirection = Trade.TradeExecutionDirection.Sell
                End If
                If tradeDirection = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item2, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandleSatisfied.Item2 + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandleSatisfied.Item2 - ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Target = .EntryPrice + ConvertFloorCeling((.EntryPrice - .Stoploss) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = GetLastDayLastCandleATR()
                    }
                ElseIf tradeDirection = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandleSatisfied.Item3, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandleSatisfied.Item3 - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = _lotSize,
                        .Stoploss = signalCandleSatisfied.Item3 + ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Target = .EntryPrice - ConvertFloorCeling((.Stoploss - .EntryPrice) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = GetLastDayLastCandleATR()
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
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))

        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signalCandle As Payload = currentTrade.SignalCandle
            If signalCandle IsNot Nothing Then
                Dim signalCandleSatisfied As Tuple(Of Boolean, Decimal, Decimal) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    Dim tradeDirection As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
                    Dim middlePoint As Decimal = (signalCandleSatisfied.Item2 + signalCandleSatisfied.Item3) / 2
                    Dim range As Decimal = signalCandleSatisfied.Item2 - middlePoint
                    If currentTick.Open >= middlePoint + range * 60 / 100 Then
                        tradeDirection = Trade.TradeExecutionDirection.Buy
                    ElseIf currentTick.Open <= middlePoint - range * 60 / 100 Then
                        tradeDirection = Trade.TradeExecutionDirection.Sell
                    End If
                    If currentTrade.EntryDirection <> tradeDirection Then
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
        If Not _entryChanged AndAlso GetLastDayLastCandleATR() <> Decimal.MinValue AndAlso
            _potentialHighEntryPrice <> Decimal.MinValue AndAlso _potentialLowEntryPrice <> Decimal.MinValue Then
            If currentTick.High >= _potentialHighEntryPrice Then
                If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.SL Then
                    Dim middlePoint As Decimal = _potentialHighEntryPrice
                    _potentialHighEntryPrice = middlePoint + ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    _potentialLowEntryPrice = middlePoint - ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                ElseIf _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Breakout Then
                    _potentialLowEntryPrice = _potentialHighEntryPrice - 2 * ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    'ElseIf _userInputs.LevelType = ATRFixedLevelBasedStrategyRuleEntities.TypeOfLevel.Candle Then
                    '    _potentialLowEntryPrice = _potentialHighEntryPrice - 2 * ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                End If
                _entryChanged = True
            ElseIf currentTick.Low <= _potentialLowEntryPrice Then
                If _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.SL Then
                    Dim middlePoint As Decimal = _potentialLowEntryPrice
                    _potentialHighEntryPrice = middlePoint + ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    _potentialLowEntryPrice = middlePoint - ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                ElseIf _userInputs.LevelType = StrategyRuleEntities.TypeOfLevel.Breakout Then
                    _potentialHighEntryPrice = _potentialLowEntryPrice + 2 * ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    'ElseIf _userInputs.LevelType = ATRFixedLevelBasedStrategyRuleEntities.TypeOfLevel.Candle Then
                    '    _potentialLowEntryPrice = _potentialHighEntryPrice - 2 * ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                End If
                _entryChanged = True
            End If
        End If
    End Function

    Private Function IsSignalCandle(ByVal candle As Payload) As Tuple(Of Boolean, Decimal, Decimal)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal) = Nothing
        If candle IsNot Nothing AndAlso Not candle.PreviousCandlePayload.DeadCandle Then
            If _potentialHighEntryPrice <> Decimal.MinValue AndAlso _potentialLowEntryPrice <> Decimal.MinValue Then
                ret = New Tuple(Of Boolean, Decimal, Decimal)(True, _potentialHighEntryPrice, _potentialLowEntryPrice)
            Else
                Dim atr As Decimal = _ATRPayload(candle.PayloadDate)
                If candle.CandleRange < atr Then
                    _potentialHighEntryPrice = candle.High
                    _potentialLowEntryPrice = candle.Low
                    _signalCandle = candle
                    ret = New Tuple(Of Boolean, Decimal, Decimal)(True, _potentialHighEntryPrice, _potentialLowEntryPrice)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetLastDayLastCandleATR() As Decimal
        Dim ret As Decimal = Decimal.MinValue
        Dim lastDayLastCandle As Payload = Nothing
        If _signalPayload IsNot Nothing AndAlso _signalPayload.Count > 0 Then
            Dim previousDayPayload As IEnumerable(Of KeyValuePair(Of Date, Payload)) = _signalPayload.Where(Function(x)
                                                                                                                Return x.Key.Date <> _tradingDate.Date
                                                                                                            End Function)
            If previousDayPayload IsNot Nothing AndAlso previousDayPayload.Count > 0 Then
                lastDayLastCandle = previousDayPayload.LastOrDefault.Value
            End If
        End If
        If lastDayLastCandle IsNot Nothing Then
            If _ATRPayload IsNot Nothing AndAlso _ATRPayload.Count > 0 Then
                ret = ConvertFloorCeling(_ATRPayload(lastDayLastCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
            End If
        End If
        Return ret
    End Function
End Class

Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MomentumReversalv2StrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public TargetMultiplier As Decimal
        Public StoplossMultiplier As Decimal
        Public ReEntryAtPreviousSignal As Boolean
        Public BreakevenMovement As Boolean
    End Class
#End Region

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
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentTick, Trade.TypeOfTrade.MIS)
            Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
            If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                If lastExecutedTrade Is Nothing Then
                    signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                Else
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedTrade.ExitTime, _signalPayload))
                    If exitCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                        signalCandle = currentMinuteCandlePayload.PreviousCandlePayload
                    End If
                End If
            ElseIf _userInputs.ReEntryAtPreviousSignal AndAlso lastExecutedTrade IsNot Nothing AndAlso
                lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso lastExecutedTrade.PLPoint < 0 Then
                signalCandle = lastExecutedTrade.SignalCandle
                signalCandleSatisfied = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, lastExecutedTrade.EntryDirection)
            End If
            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                If signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.High + buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Buy,
                        .Quantity = LotSize,
                        .Stoploss = .EntryPrice - ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Target = .EntryPrice + ConvertFloorCeling(_ATRPayload(signalCandle.PayloadDate) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = ConvertFloorCeling(GetLastDayLastCandleATR(), _parentStrategy.TickSize, RoundOfType.Celing),
                        .Supporting3 = _ATRPayload(signalCandle.PayloadDate)
                    }
                ElseIf signalCandleSatisfied.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                        .EntryPrice = signalCandle.Low - buffer,
                        .EntryDirection = Trade.TradeExecutionDirection.Sell,
                        .Quantity = LotSize,
                        .Stoploss = .EntryPrice + ConvertFloorCeling(GetLastDayLastCandleATR() * _userInputs.StoplossMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Target = .EntryPrice - ConvertFloorCeling(_ATRPayload(signalCandle.PayloadDate) * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing),
                        .Buffer = buffer,
                        .SignalCandle = signalCandle,
                        .OrderType = Trade.TypeOfOrder.SL,
                        .Supporting1 = signalCandle.PayloadDate.ToShortTimeString,
                        .Supporting2 = ConvertFloorCeling(GetLastDayLastCandleATR(), _parentStrategy.TickSize, RoundOfType.Celing),
                        .Supporting3 = _ATRPayload(signalCandle.PayloadDate)
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
                Dim signalCandleSatisfied As Tuple(Of Boolean, Trade.TradeExecutionDirection) = IsSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload)
                If signalCandleSatisfied IsNot Nothing AndAlso signalCandleSatisfied.Item1 Then
                    If signalCandle.PayloadDate <> currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate Then
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
                Dim excpectedTarget As Decimal = currentTrade.EntryPrice + (currentTrade.PotentialTarget - currentTrade.EntryPrice) * 2 / 3
                If currentTick.Open >= excpectedTarget Then
                    triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, LotSize, _parentStrategy.StockType)
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim excpectedTarget As Decimal = currentTrade.EntryPrice - (currentTrade.EntryPrice - currentTrade.PotentialTarget) * 2 / 3
                If currentTick.Open <= excpectedTarget Then
                    triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, LotSize, _parentStrategy.StockType)
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

    Private Function IsSignalCandle(ByVal candle As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso
            Not candle.DojiCandle AndAlso Not candle.PreviousCandlePayload.DojiCandle Then
            Dim previousCandleSatisfied As Boolean = IsCandleCloseSatisfied(candle.PreviousCandlePayload, 1)
            Dim currentCandleSatisfied As Boolean = False
            If Not previousCandleSatisfied Then
                previousCandleSatisfied = IsCandleCloseSatisfied(candle.PreviousCandlePayload, -1)
                If previousCandleSatisfied Then
                    currentCandleSatisfied = IsCandleCloseSatisfied(candle, 1)
                    If currentCandleSatisfied Then ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Buy)
                End If
            Else
                currentCandleSatisfied = IsCandleCloseSatisfied(candle, -1)
                If currentCandleSatisfied Then ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection)(True, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function

    Private Function IsCandleCloseSatisfied(ByVal candle As Payload, ByVal side As Integer) As Boolean
        Dim ret As Boolean = False
        If candle IsNot Nothing Then
            Dim candleCloseStrength As Decimal = ConvertFloorCeling(candle.CandleRange * 25 / 100, _parentStrategy.TickSize, RoundOfType.Celing)
            If side = 1 Then
                If (candle.High - candle.Close) * 4 <= candle.CandleRange Then
                    ret = True
                End If
            ElseIf side = -1 Then
                If (candle.Close - candle.Low) * 4 <= candle.CandleRange Then
                    ret = True
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
                ret = _ATRPayload(lastDayLastCandle.PayloadDate)
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function
End Class

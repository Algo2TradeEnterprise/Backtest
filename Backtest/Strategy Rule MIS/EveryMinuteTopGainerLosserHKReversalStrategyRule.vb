Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation
Imports Utilities.Numbers

Public Class EveryMinuteTopGainerLosserHKReversalStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public TargetMultiplier As Decimal
        Public BreakevenMovement As Boolean
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private ReadOnly _userInputs As StrategyRuleEntities
    Private ReadOnly _signalTime As Date
    Private ReadOnly _gainLossPer As Decimal

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal signalTime As Date,
                   ByVal gainLossPer As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _signalTime = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, signalTime.Hour, signalTime.Minute, signalTime.Second)
        _gainLossPer = gainLossPer
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso currentMinuteCandlePayload.PayloadDate > _signalTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                signalCandle = signal.Item2
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                    If entryPrice = signalCandle.High Then entryPrice = entryPrice + buffer
                    Dim lowestPointOfTheDay As Decimal = GetLowestPointOfTheDay(signalCandle)
                    Dim stoplossPrice As Decimal = ConvertFloorCeling(lowestPointOfTheDay, _parentStrategy.TickSize, RoundOfType.Floor)
                    If stoplossPrice = lowestPointOfTheDay Then stoplossPrice = stoplossPrice - buffer
                    Dim slPoint As Decimal = entryPrice - stoplossPrice
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)

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
                                    .Supporting2 = slPoint,
                                    .Supporting3 = _signalTime.ToString("HH:mm:ss"),
                                    .Supporting4 = _gainLossPer
                                }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(entryPrice, RoundOfType.Floor)
                    If entryPrice = signalCandle.Low Then entryPrice = entryPrice - buffer
                    Dim highestPointOfTheDay As Decimal = GetHighestPointOfTheDay(signalCandle)
                    Dim stoplossPrice As Decimal = ConvertFloorCeling(highestPointOfTheDay, _parentStrategy.TickSize, RoundOfType.Celing)
                    If stoplossPrice = highestPointOfTheDay Then stoplossPrice = stoplossPrice + buffer
                    Dim slPoint As Decimal = stoplossPrice - entryPrice
                    Dim targetPoint As Decimal = ConvertFloorCeling(slPoint * _userInputs.TargetMultiplier, _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, _userInputs.MaxLossPerTrade, _parentStrategy.StockType)

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
                                    .Supporting2 = slPoint,
                                    .Supporting3 = _signalTime.ToString("HH:mm:ss"),
                                    .Supporting4 = _gainLossPer
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
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.BreakevenMovement AndAlso currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim slPoint As Decimal = currentTrade.Supporting2
            Dim triggerPrice As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open >= currentTrade.EntryPrice + slPoint Then
                    triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Buy, Me.LotSize, _parentStrategy.StockType)
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open <= currentTrade.EntryPrice - slPoint Then
                    triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, Trade.TradeExecutionDirection.Sell, Me.LotSize, _parentStrategy.StockType)
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Breakeven at {1}", slPoint, currentTick.PayloadDate.ToString("HH:mm:ss")))
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            If _gainLossPer < 0 AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, hkCandle, Trade.TradeExecutionDirection.Buy)
            ElseIf _gainLossPer > 0 AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection)(True, hkCandle, Trade.TradeExecutionDirection.Sell)
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestPointOfTheDay(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _hkPayload.Max(Function(x)
                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                     Return x.Value.High
                                 Else
                                     Return Decimal.MinValue
                                 End If
                             End Function)
        Return ret
    End Function

    Private Function GetLowestPointOfTheDay(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _hkPayload.Min(Function(x)
                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                     Return x.Value.Low
                                 Else
                                     Return Decimal.MaxValue
                                 End If
                             End Function)
        Return ret
    End Function

#Region "Stock Selection"
    Public Class InstumentDetails
        Public TradingSymbol As String
        Public LotSize As Integer
        Public SignalTime As Date
        Public GainLossPer As Decimal
        Public SignalTriggerTime As Date
    End Class
    Public Shared Function GetFirstTradeTriggerTime(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tradingDate As Date, ByVal signalTime As Date, ByVal gainLossPer As Decimal) As Date
        Dim ret As Date = Date.MaxValue
        Dim hkPayload As Dictionary(Of Date, Payload) = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(inputPayload, hkPayload)
        Dim stockSignalTime As Date = New Date(tradingDate.Year, tradingDate.Month, tradingDate.Day, signalTime.Hour, signalTime.Minute, signalTime.Second)
        Dim triggerPrice As Decimal = Decimal.MinValue
        For Each runningPayload In inputPayload.OrderBy(Function(x)
                                                            Return x.Key
                                                        End Function)
            If runningPayload.Key > stockSignalTime AndAlso runningPayload.Key.Date = tradingDate.Date AndAlso
                runningPayload.Value.PreviousCandlePayload.PayloadDate.Date = tradingDate.Date Then
                Dim hkCandle As Payload = hkPayload(runningPayload.Value.PreviousCandlePayload.PayloadDate)
                If gainLossPer < 0 AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    triggerPrice = ConvertFloorCeling(hkCandle.High, 0.05, RoundOfType.Celing)
                    Dim buffer As Decimal = CalculateDuplicateBuffer(triggerPrice, RoundOfType.Floor)
                    If triggerPrice = hkCandle.High Then triggerPrice = triggerPrice + buffer
                ElseIf gainLossPer > 0 AndAlso hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                    triggerPrice = ConvertFloorCeling(hkCandle.Low, 0.05, RoundOfType.Floor)
                    Dim buffer As Decimal = CalculateDuplicateBuffer(triggerPrice, RoundOfType.Floor)
                    If triggerPrice = hkCandle.Low Then triggerPrice = triggerPrice - buffer
                End If
                If triggerPrice <> Decimal.MinValue Then
                    If gainLossPer < 0 Then
                        If runningPayload.Value.High >= triggerPrice Then
                            For Each runningTick In runningPayload.Value.Ticks
                                If runningTick.High >= triggerPrice Then
                                    ret = runningTick.PayloadDate
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf gainLossPer > 0 Then
                        If runningPayload.Value.Low <= triggerPrice Then
                            For Each runningTick In runningPayload.Value.Ticks
                                If runningTick.Low <= triggerPrice Then
                                    ret = runningTick.PayloadDate
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    If ret <> Date.MaxValue Then Exit For
                End If
            End If
        Next
        Return ret
    End Function
    Private Shared Function CalculateDuplicateBuffer(ByVal price As Decimal, ByVal floorOrCeiling As RoundOfType) As Decimal
        Dim bufferPrice As Decimal = Nothing
        'Assuming 1% target, we can afford to have buffer as 2.5% of that 1% target
        bufferPrice = NumberManipulation.ConvertFloorCeling(price * 0.01 * 0.025, 0.05, floorOrCeiling)
        Return bufferPrice
    End Function
#End Region

End Class
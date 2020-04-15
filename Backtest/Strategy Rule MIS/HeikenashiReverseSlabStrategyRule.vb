Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class HeikenashiReverseSlabStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public SatelliteTradeTargetMultiplier As Decimal
    End Class
#End Region

    Private _hkPayload As Dictionary(Of Date, Payload) = Nothing
    Private ReadOnly _slab As Decimal
    Private ReadOnly _userInputs As StrategyRuleEntities
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal slab As Decimal)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        _slab = slab
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim anchorTrade As Trade = GetMainAnchorTrade(currentMinuteCandlePayload)
            If Not IsLogicalActiveTrade(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Buy) Then
                If (Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) OrElse anchorTrade IsNot Nothing) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                    Dim signal As Tuple(Of Boolean, Payload, Decimal) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                    If signal IsNot Nothing AndAlso signal.Item1 Then
                        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Buy)
                        If lastExecutedOrder Is Nothing OrElse lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item3, RoundOfType.Floor)
                            Dim entryPrice As Decimal = signal.Item3 + buffer
                            Dim stoploss As Decimal = signal.Item3 - _slab - buffer
                            Dim quantity As Integer = Me.LotSize
                            Dim targetPoint As Decimal = 100000000000
                            Dim targetRemark As String = "Anchor"
                            If anchorTrade IsNot Nothing Then
                                targetPoint = (entryPrice - stoploss) * _userInputs.SatelliteTradeTargetMultiplier
                                targetRemark = "Satelite"
                            End If

                            If currentTick.Open < entryPrice Then
                                parameter1 = New PlaceOrderParameters With {
                                            .EntryPrice = entryPrice,
                                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                            .Quantity = quantity,
                                            .Stoploss = stoploss,
                                            .Target = entryPrice + targetPoint,
                                            .Buffer = buffer,
                                            .SignalCandle = signal.Item2,
                                            .OrderType = Trade.TypeOfOrder.SL,
                                            .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                            .Supporting2 = targetRemark,
                                            .Supporting3 = _slab
                                        }
                            End If
                        End If
                    End If
                End If
            End If

            If Not IsLogicalActiveTrade(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Sell) Then
                If (Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) OrElse anchorTrade IsNot Nothing) AndAlso
                    Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then
                    Dim signal As Tuple(Of Boolean, Payload, Decimal) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                    If signal IsNot Nothing AndAlso signal.Item1 Then
                        Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType, Trade.TradeExecutionDirection.Sell)
                        If lastExecutedOrder Is Nothing OrElse lastExecutedOrder.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item3, RoundOfType.Floor)
                            Dim entryPrice As Decimal = signal.Item3 - buffer
                            Dim stoploss As Decimal = signal.Item3 + _slab + buffer
                            Dim quantity As Integer = Me.LotSize
                            Dim targetPoint As Decimal = 100000000000
                            Dim targetRemark As String = "Anchor"
                            If anchorTrade IsNot Nothing Then
                                targetPoint = (stoploss - entryPrice) * _userInputs.SatelliteTradeTargetMultiplier
                                targetRemark = "Satelite"
                            End If

                            If currentTick.Open > entryPrice Then
                                parameter2 = New PlaceOrderParameters With {
                                            .EntryPrice = entryPrice,
                                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                            .Quantity = quantity,
                                            .Stoploss = stoploss,
                                            .Target = entryPrice - targetPoint,
                                            .Buffer = buffer,
                                            .SignalCandle = signal.Item2,
                                            .OrderType = Trade.TypeOfOrder.SL,
                                            .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                            .Supporting2 = targetRemark,
                                            .Supporting3 = _slab
                                        }
                            End If
                        End If
                    End If
                End If
            End If
        End If

        Dim parameterList As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter1)
        End If
        If parameter2 IsNot Nothing Then
            If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
            parameterList.Add(parameter2)
        End If
        If parameterList IsNot Nothing AndAlso parameterList.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameterList)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Payload, Decimal) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
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

    Private Function GetSlabBasedLevel(ByVal price As Decimal, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If direction = Trade.TradeExecutionDirection.Buy Then
            ret = Math.Ceiling(price / _slab) * _slab
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            ret = Math.Floor(price / _slab) * _slab
        End If
        Return ret
    End Function

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload, Decimal)
        Dim ret As Tuple(Of Boolean, Payload, Decimal) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            If direction = Trade.TradeExecutionDirection.Buy Then
                If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    Dim buyLevel As Decimal = GetSlabBasedLevel(hkCandle.High, Trade.TradeExecutionDirection.Buy)
                    ret = New Tuple(Of Boolean, Payload, Decimal)(True, hkCandle, buyLevel)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                    Dim sellLevel As Decimal = GetSlabBasedLevel(hkCandle.Low, Trade.TradeExecutionDirection.Sell)
                    ret = New Tuple(Of Boolean, Payload, Decimal)(True, hkCandle, sellLevel)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetAnchorTrades(ByVal candle As Payload) As List(Of Trade)
        Dim ret As List(Of Trade) = Nothing
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then
            ret = inprogressTrades.FindAll(Function(x)
                                               Return x.Supporting2 = "Anchor"
                                           End Function)
        End If
        Return ret
    End Function

    Private Function GetMainAnchorTrade(ByVal candle As Payload) As Trade
        Dim ret As Trade = Nothing
        Dim anchorTrades As List(Of Trade) = GetAnchorTrades(candle)
        If anchorTrades IsNot Nothing AndAlso anchorTrades.Count = 2 Then
            ret = anchorTrades.OrderBy(Function(x)
                                           Return x.EntryTime
                                       End Function).FirstOrDefault
        End If
        Return ret
    End Function

    Private Function IsLogicalActiveTrade(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Boolean
        Dim ret As Boolean = Nothing
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        If inprogressTrades IsNot Nothing AndAlso inprogressTrades.Count > 0 Then
            Dim directionTrades As List(Of Trade) = inprogressTrades.FindAll(Function(x)
                                                                                 Return x.EntryDirection = direction
                                                                             End Function)
            ret = directionTrades IsNot Nothing AndAlso directionTrades.Count = 2
        End If
        Return ret
    End Function
End Class
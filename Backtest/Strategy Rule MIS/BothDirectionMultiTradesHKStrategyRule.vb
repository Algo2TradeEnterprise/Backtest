Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class BothDirectionMultiTradesHKStrategyRule
    Inherits StrategyRule

    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _quantity As Integer = Integer.MinValue

    Private ReadOnly _maxSingleDirectionTradeCount As Integer = 3
    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.HeikenAshi.ConvertToHeikenAshi(_signalPayload, _hkPayload)
        Indicator.ATR.CalculateATR(14, _hkPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameterList As List(Of PlaceOrderParameters) = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim buyActiveTrades As Integer = GetActiveTradeCount(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Buy)
            Dim sellActiveTrades As Integer = GetActiveTradeCount(currentMinuteCandlePayload, Trade.TradeExecutionDirection.Sell)
            If buyActiveTrades < _maxSingleDirectionTradeCount Then
                Dim signal As Tuple(Of Boolean, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Buy)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = GetPriceForHK(signal.Item2, Trade.TradeExecutionDirection.Buy)
                    Dim stoploss As Decimal = GetPriceForHK(signal.Item2, Trade.TradeExecutionDirection.Sell)
                    Dim quantity As Integer = GetQuantity(signal.Item2)
                    For tradeCtr As Integer = 1 To _maxSingleDirectionTradeCount - buyActiveTrades
                        Dim slPrice As Decimal = stoploss
                        If tradeCtr > 1 Then slPrice = entryPrice - 10000000
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                   .EntryPrice = entryPrice,
                                                                   .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                                   .Quantity = quantity,
                                                                   .Stoploss = slPrice,
                                                                   .Target = .EntryPrice + 10000000,
                                                                   .Buffer = buffer,
                                                                   .SignalCandle = signal.Item2,
                                                                   .OrderType = Trade.TypeOfOrder.SL,
                                                                   .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                                                   .Supporting2 = tradeCtr
                                                               }

                        If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
                        parameterList.Add(parameter)
                    Next
                End If
            End If
            If sellActiveTrades < _maxSingleDirectionTradeCount Then
                Dim signal As Tuple(Of Boolean, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, Trade.TradeExecutionDirection.Sell)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2.Open, RoundOfType.Floor)
                    Dim entryPrice As Decimal = GetPriceForHK(signal.Item2, Trade.TradeExecutionDirection.Sell)
                    Dim stoploss As Decimal = GetPriceForHK(signal.Item2, Trade.TradeExecutionDirection.Buy)
                    Dim quantity As Integer = GetQuantity(signal.Item2)
                    For tradeCtr As Integer = 1 To _maxSingleDirectionTradeCount - sellActiveTrades
                        Dim slPrice As Decimal = stoploss
                        If tradeCtr > 1 Then slPrice = entryPrice + 10000000
                        Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                                   .EntryPrice = entryPrice,
                                                                   .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                                   .Quantity = quantity,
                                                                   .Stoploss = slPrice,
                                                                   .Target = .EntryPrice - 10000000,
                                                                   .Buffer = buffer,
                                                                   .SignalCandle = signal.Item2,
                                                                   .OrderType = Trade.TypeOfOrder.SL,
                                                                   .Supporting1 = signal.Item2.PayloadDate.ToString("HH:mm:ss"),
                                                                   .Supporting2 = tradeCtr
                                                               }

                        If parameterList Is Nothing Then parameterList = New List(Of PlaceOrderParameters)
                        parameterList.Add(parameter)
                    Next
                End If
            End If
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
            Dim signal As Tuple(Of Boolean, Payload) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick, currentTrade.EntryDirection)
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            Dim lastExitTrade As Trade = _parentStrategy.GetLastExitTradeOfTheStock(candle, Trade.TypeOfTrade.MIS, direction)
            If direction = Trade.TradeExecutionDirection.Buy AndAlso
                hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                If lastExitTrade IsNot Nothing Then
                    Dim buyLevel As Decimal = GetPriceForHK(hkCandle, Trade.TradeExecutionDirection.Buy)
                    If lastExitTrade.ExitPrice - buyLevel >= Math.Abs(lastExitTrade.PLPoint) * 1.5 Then
                        ret = New Tuple(Of Boolean, Payload)(True, hkCandle)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Payload)(True, hkCandle)
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell AndAlso
                hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bullish Then
                If lastExitTrade IsNot Nothing Then
                    Dim sellLevel As Decimal = GetPriceForHK(hkCandle, Trade.TradeExecutionDirection.Sell)
                    If sellLevel - lastExitTrade.ExitPrice >= Math.Abs(lastExitTrade.PLPoint) * 1.5 Then
                        ret = New Tuple(Of Boolean, Payload)(True, hkCandle)
                    End If
                Else
                    ret = New Tuple(Of Boolean, Payload)(True, hkCandle)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetActiveTradeCount(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Integer
        Dim ret As Integer = 0
        Dim openTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Open)
        Dim inprogressTrades As List(Of Trade) = _parentStrategy.GetSpecificTrades(candle, Trade.TypeOfTrade.MIS, Trade.TradeExecutionStatus.Inprogress)
        Dim allTrades As List(Of Trade) = New List(Of Trade)
        If openTrades IsNot Nothing Then allTrades.AddRange(openTrades)
        If inprogressTrades IsNot Nothing Then allTrades.AddRange(inprogressTrades)
        If allTrades IsNot Nothing AndAlso allTrades.Count > 0 Then
            Dim directionTrades As List(Of Trade) = allTrades.FindAll(Function(x)
                                                                          Return x.EntryDirection = direction
                                                                      End Function)
            If directionTrades IsNot Nothing AndAlso directionTrades.Count > 0 Then
                ret = directionTrades.Count
            End If
        End If
        Return ret
    End Function

    Private Function GetPriceForHK(ByVal hkCandle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If hkCandle IsNot Nothing Then
            If direction = Trade.TradeExecutionDirection.Buy Then
                ret = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                If ret = Math.Round(hkCandle.High, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(ret, RoundOfType.Floor)
                    ret = ret + buffer
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                ret = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                If ret = Math.Round(hkCandle.Low, 2) Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(ret, RoundOfType.Floor)
                    ret = ret - buffer
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATROfTheDay(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _atrPayload.Max(Function(x)
                                  If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                      Return x.Value
                                  Else
                                      Return Decimal.MinValue
                                  End If
                              End Function)
        Return ret
    End Function

    Private Function GetQuantity(ByVal signalCandle As Payload) As Integer
        If _quantity = Integer.MinValue Then
            Dim highestATR As Decimal = ConvertFloorCeling(GetHighestATROfTheDay(signalCandle), _parentStrategy.TickSize, RoundOfType.Floor)
            _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, signalCandle.Open, signalCandle.Open + highestATR, 500, Trade.TypeOfStock.Cash)
            Return _quantity
        Else
            Return _quantity
        End If
    End Function
End Class
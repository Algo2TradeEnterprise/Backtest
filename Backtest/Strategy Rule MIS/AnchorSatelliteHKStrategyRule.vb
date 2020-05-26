Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class AnchorSatelliteHKStrategyRule
    Inherits StrategyRule

    Private _swingHighPayload As Dictionary(Of Date, Decimal)
    Private _swingLowPayload As Dictionary(Of Date, Decimal)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _quantity As Integer = Integer.MinValue

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
        Indicator.SwingHighLow.CalculateSwingHighLow(_hkPayload, True, _swingHighPayload, _swingLowPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
                If lastExecutedOrder Is Nothing Then
                    signalCandle = signal.Item3
                ElseIf lastExecutedOrder IsNot Nothing Then
                    If signal.Item3.PayloadDate <> lastExecutedOrder.SignalCandle.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                End If
            End If
            If signalCandle IsNot Nothing Then
                If _quantity = Integer.MinValue Then
                    Dim targetPoint As Decimal = ConvertFloorCeling(GetHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                    _quantity = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, currentTick.Open, currentTick.Open + targetPoint, _parentStrategy.StockMaxProfitPerDay, Trade.TypeOfStock.Cash)
                End If
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) AndAlso
                        Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Buy) Then
                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                        If entryPrice = signalCandle.High Then entryPrice = entryPrice + buffer
                        Dim slPrice As Decimal = ConvertFloorCeling(signalCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                        If slPrice = signalCandle.Low Then slPrice = slPrice - buffer
                        Dim quantity As Integer = _quantity

                        parameter = New PlaceOrderParameters With {
                            .EntryPrice = entryPrice,
                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                            .Quantity = quantity,
                            .Stoploss = slPrice,
                            .Target = .EntryPrice + 10000000,
                            .Buffer = buffer,
                            .SignalCandle = signalCandle,
                            .OrderType = Trade.TypeOfOrder.SL,
                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                        }
                    End If
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    If Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) AndAlso
                        Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS, Trade.TradeExecutionDirection.Sell) Then

                        Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                        Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                        If entryPrice = signalCandle.Low Then entryPrice = entryPrice - buffer
                        Dim slPrice As Decimal = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                        If slPrice = signalCandle.High Then slPrice = slPrice + buffer
                        Dim quantity As Integer = _quantity

                        parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = slPrice,
                                .Target = .EntryPrice - 10000000,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                            }
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter IsNot Nothing Then
            If parameters Is Nothing Then parameters = New List(Of PlaceOrderParameters)
            parameters.Add(parameter)
        End If
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetSignalCandle(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection = signal.Item4 AndAlso currentTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentMinuteCandlePayload.PayloadDate <= currentTrade.EntryTime Then
                Dim triggerPrice As Decimal = Decimal.MinValue
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                    'Dim lowestPoint As Decimal = GetLowestPointOfTheDay(currentMinuteCandlePayload.PreviousCandlePayload)
                    Dim lowestPoint As Decimal = Math.Min(_swingLowPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), currentTrade.SignalCandle.Low)
                    triggerPrice = ConvertFloorCeling(lowestPoint, _parentStrategy.TickSize, RoundOfType.Floor)
                    If triggerPrice = lowestPoint Then triggerPrice = triggerPrice - currentTrade.StoplossBuffer
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                    'Dim highestPoint As Decimal = GetHighestPointOfTheDay(currentMinuteCandlePayload.PreviousCandlePayload)
                    Dim highestPoint As Decimal = Math.Max(_swingHighPayload(currentMinuteCandlePayload.PreviousCandlePayload.PayloadDate), currentTrade.SignalCandle.High)
                    triggerPrice = ConvertFloorCeling(highestPoint, _parentStrategy.TickSize, RoundOfType.Celing)
                    If triggerPrice = highestPoint Then triggerPrice = triggerPrice + currentTrade.StoplossBuffer
                End If
                If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, Math.Abs(Math.Round(triggerPrice - currentTrade.EntryPrice)))
                End If
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

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If candle IsNot Nothing AndAlso Not candle.DeadCandle Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            If Math.Round(hkCandle.Open, 4) = Math.Round(hkCandle.High, 4) Then
                If hkCandle.Low < hkCandle.PreviousCandlePayload.Low Then
                    Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf Math.Round(hkCandle.Open, 4) = Math.Round(hkCandle.Low, 4) Then
                If hkCandle.High > hkCandle.PreviousCandlePayload.High Then
                    Dim entryPrice As Decimal = ConvertFloorCeling(hkCandle.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, entryPrice, hkCandle, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 Then
            Dim firstCandle As Payload = Nothing
            For Each runningPayload In _signalPayload
                If runningPayload.Key.Date = _tradingDate.Date Then
                    If runningPayload.Value.PreviousCandlePayload.PayloadDate.Date <> _tradingDate.Date Then
                        firstCandle = runningPayload.Value
                        Exit For
                    End If
                End If
            Next
            If firstCandle IsNot Nothing AndAlso firstCandle.PreviousCandlePayload IsNot Nothing Then
                ret = _atrPayload.Max(Function(x)
                                          If x.Key.Date >= firstCandle.PreviousCandlePayload.PayloadDate.Date AndAlso
                                            x.Key <= signalCandle.PayloadDate Then
                                              Return x.Value
                                          Else
                                              Return Decimal.MinValue
                                          End If
                                      End Function)
            End If
        End If
        Return ret
    End Function

    Private Function GetHighestPointOfTheDay(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        ret = _signalPayload.Max(Function(x)
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
        ret = _signalPayload.Min(Function(x)
                                     If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                         Return x.Value.Low
                                     Else
                                         Return Decimal.MaxValue
                                     End If
                                 End Function)
        Return ret
    End Function
End Class
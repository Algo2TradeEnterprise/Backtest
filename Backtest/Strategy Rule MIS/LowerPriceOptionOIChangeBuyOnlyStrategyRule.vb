Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowerPriceOptionOIChangeBuyOnlyStrategyRule
    Inherits StrategyRule

    Private _hkPayload As Dictionary(Of Date, Payload)
    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _targetPoint As Decimal

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
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandlePayload.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim quantity As Integer = Integer.MinValue
            Dim orderType As Trade.TypeOfOrder = Trade.TypeOfOrder.SL
            Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, Trade.TypeOfTrade.MIS)
            If lastExecutedOrder IsNot Nothing Then
                Dim averagePrice As Decimal = lastExecutedOrder.Supporting2
                If currentTick.Open <= averagePrice - _targetPoint Then
                    signalCandle = currentMinuteCandlePayload
                    quantity = lastExecutedOrder.Quantity * 2
                End If
            Else
                Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    signalCandle = signal.Item2
                    _targetPoint = ConvertFloorCeling(_atrPayload(signalCandle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Celing)
                    Dim expectedQuantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, currentTick.Open, currentTick.Open + _targetPoint, 500, Trade.TypeOfStock.Futures)
                End If
            End If
            If signalCandle IsNot Nothing AndAlso quantity <> Integer.MinValue Then
                Dim entryPrice As Decimal = ConvertFloorCeling(signalCandle.High, _parentStrategy.TickSize, RoundOfType.Celing)

                parameter = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = 0,
                                    .Target = .EntryPrice + _targetPoint,
                                    .Buffer = 0,
                                    .SignalCandle = signalCandle,
                                    .OrderType = orderType,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                                }
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter IsNot Nothing Then
            parameters = New List(Of PlaceOrderParameters) From {
                parameter
            }
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
            Dim signal As Tuple(Of Boolean, Payload) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload)
        Dim ret As Tuple(Of Boolean, Payload) = Nothing
        If candle IsNot Nothing Then
            Dim hkCandle As Payload = _hkPayload(candle.PayloadDate)
            If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                ret = New Tuple(Of Boolean, Payload)(True, hkCandle)
            End If
        End If
        Return ret
    End Function

#Region "Stock Selection"
    Public Class OptionInstumentDetails
        Public TradingSymbol As String
        Public LotSize As Integer
        Public InstrumentType As String
        Public TriggerTime As Date
    End Class
    Public Shared Function GetFirstTradeTriggerTime(ByVal inputPayload As Dictionary(Of Date, Payload), ByVal tradingDate As Date) As Date
        Dim ret As Date = Date.MinValue
        Dim hkPayload As Dictionary(Of Date, Payload) = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(inputPayload, hkPayload)
        Dim triggerPrice As Decimal = Decimal.MinValue
        For Each runningPayload In inputPayload.OrderBy(Function(x)
                                                            Return x.Key
                                                        End Function)
            If runningPayload.Key.Date = tradingDate.Date AndAlso
                runningPayload.Value.PreviousCandlePayload.PayloadDate.Date = tradingDate.Date Then
                Dim hkCandle As Payload = hkPayload(runningPayload.Key)
                If hkCandle.CandleStrengthHeikenAshi = Payload.StrongCandle.Bearish Then
                    triggerPrice = ConvertFloorCeling(hkCandle.High, 0.05, RoundOfType.Celing)
                End If
                If triggerPrice <> Decimal.MinValue Then
                    For Each runningTick In runningPayload.Value.Ticks
                        If runningTick.High >= triggerPrice Then
                            ret = runningTick.PayloadDate
                            Exit For
                        End If
                    Next
                    If ret <> Date.MinValue Then Exit For
                End If
            End If
        Next
        Return ret
    End Function
#End Region

End Class
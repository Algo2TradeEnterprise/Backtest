Imports Algo2TradeBLL
Imports System.Threading
Imports Backtest.StrategyHelper
Imports Utilities.Numbers.NumberManipulation

Public Class EMAAttractionStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public EMAPeriod As Integer
        Public MaxLossPerTrade As Decimal
        Public ATRMultipler As Decimal
        Public TakeDoubleQuantity As Boolean
    End Class
#End Region

    Private ReadOnly _userInputs As StrategyRuleEntities

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _emaPayload As Dictionary(Of Date, Decimal)

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.EMA.CalculateEMA(_userInputs.EMAPeriod, Payload.PayloadFields.Close, _signalPayload, _emaPayload)
        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(ByVal currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter1 As PlaceOrderParameters = Nothing
        Dim parameter2 As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentMinuteCandle, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetSignalCandle(currentMinuteCandle.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
                If lastExecutedOrder IsNot Nothing Then
                    If currentMinuteCandle.PayloadDate > lastExecutedOrder.ExitTime Then
                        signalCandle = signal.Item4
                    End If
                Else
                    signalCandle = signal.Item4
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                Dim entryPrice As Decimal = signal.Item2
                Dim stoplossPrice As Decimal = signal.Item3
                Dim atr As Decimal = Math.Round(_atrPayload(signalCandle.PayloadDate), 4)
                Dim ema As Decimal = Math.Round(_emaPayload(signalCandle.PayloadDate), 4)

                If signal.Item5 = Trade.TradeExecutionDirection.Buy Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoplossPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = .EntryPrice + 10000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = atr,
                                    .Supporting3 = ema,
                                    .Supporting4 = 1,
                                    .Supporting5 = signal.Item6
                                }

                    If _userInputs.TakeDoubleQuantity Then
                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = .EntryPrice + 10000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = atr,
                                    .Supporting3 = ema,
                                    .Supporting4 = 2,
                                    .Supporting5 = signal.Item6
                                }
                    End If
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell Then
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoplossPrice, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, _parentStrategy.StockType)

                    parameter1 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = .EntryPrice - 10000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = atr,
                                    .Supporting3 = ema,
                                    .Supporting4 = 1,
                                    .Supporting5 = signal.Item6
                                }

                    If _userInputs.TakeDoubleQuantity Then
                        parameter2 = New PlaceOrderParameters With {
                                    .EntryPrice = entryPrice,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = quantity,
                                    .Stoploss = stoplossPrice,
                                    .Target = .EntryPrice - 10000000,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = atr,
                                    .Supporting3 = ema,
                                    .Supporting4 = 2,
                                    .Supporting5 = signal.Item6
                                }
                    End If
                End If
            End If
        End If
        Dim parameters As List(Of PlaceOrderParameters) = Nothing
        If parameter1 IsNot Nothing Then parameters = New List(Of PlaceOrderParameters) From {parameter1}
        If parameter2 IsNot Nothing Then parameters.Add(parameter2)
        If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
            ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, parameters)
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String) = GetSignalCandle(currentMinuteCandle.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        ElseIf currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso currentTrade.Supporting4 = 1 Then
            Dim ema As Decimal = _emaPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open >= ema Then
                    ret = New Tuple(Of Boolean, String)(True, "EMA Target")
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open <= ema Then
                    ret = New Tuple(Of Boolean, String)(True, "EMA Target")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If _userInputs.TakeDoubleQuantity AndAlso currentTrade IsNot Nothing AndAlso
            currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress AndAlso currentTrade.Supporting4 = 2 Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim triggerPrice As Decimal = Decimal.MinValue
            Dim slRemark As Decimal = 0
            Dim ema As Decimal = _emaPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentTick.Open >= ema Then
                    triggerPrice = currentTrade.EntryPrice + _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                    slRemark = currentTrade.EntryPrice - triggerPrice
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentTick.Open <= ema Then
                    triggerPrice = currentTrade.EntryPrice - _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
                    slRemark = triggerPrice - currentTrade.EntryPrice
                End If
            End If
            If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, slRemark)
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(ByVal currentTick As Payload, ByVal currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Private Function GetSignalCandle(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String) = Nothing
        If candle IsNot Nothing Then
            Dim atr As Decimal = ConvertFloorCeling(_atrPayload(candle.PayloadDate) * _userInputs.ATRMultipler, _parentStrategy.TickSize, RoundOfType.Celing)
            Dim ema As Decimal = _emaPayload(candle.PayloadDate)
            If candle.High < ema Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
                Dim entryPrice As Decimal = candle.High + buffer
                Dim diff As Decimal = ConvertFloorCeling(ema - entryPrice, _parentStrategy.TickSize, RoundOfType.Celing)
                If diff > atr Then
                    Dim slPoint As Decimal = candle.CandleRange + 2 * buffer
                    Dim slRemark As String = "CR"
                    If slPoint >= diff Then
                        slPoint = atr
                        slRemark = "ATR"
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, entryPrice, entryPrice - slPoint, candle, Trade.TradeExecutionDirection.Buy, slRemark)
                End If
            ElseIf candle.Low > ema Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
                Dim entryPrice As Decimal = candle.Low - buffer
                Dim diff As Decimal = ConvertFloorCeling(entryPrice - ema, _parentStrategy.TickSize, RoundOfType.Celing)
                If diff > atr Then
                    Dim slPoint As Decimal = candle.CandleRange + 2 * buffer
                    Dim slRemark As String = "CR"
                    If slPoint >= diff Then
                        slPoint = atr
                        slRemark = "ATR"
                    End If
                    ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection, String)(True, entryPrice, entryPrice + slPoint, candle, Trade.TradeExecutionDirection.Sell, slRemark)
                End If
            End If
        End If
        Return ret
    End Function
End Class
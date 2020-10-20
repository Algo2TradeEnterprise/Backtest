Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class MADirectionBasedStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public ATRMultiplier As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerHighPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _bollingerLowPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing

    Private ReadOnly _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
    Private ReadOnly _gap As Boolean = False

    Private ReadOnly _userInputs As StrategyRuleEntities

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities,
                   ByVal direction As Integer,
                   ByVal gap As Integer)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = _entities
        If direction > 0 Then
            _direction = Trade.TradeExecutionDirection.Buy
        ElseIf direction < 0 Then
            _direction = Trade.TradeExecutionDirection.Sell
        End If
        _gap = gap
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.BollingerBands.CalculateBollingerBands(20, Payload.PayloadFields.Close, 2, _signalPayload, _bollingerHighPayload, _bollingerLowPayload, _smaPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastOrder IsNot Nothing Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastOrder.ExitTime, _signalPayload))
                    If currentMinuteCandle.PayloadDate > exitCandle.PayloadDate Then
                        signalCandle = signal.Item3
                    End If
                Else
                    signalCandle = signal.Item3
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim entryPrice As Decimal = signal.Item2
                If signal.Item4 = Trade.TradeExecutionDirection.Buy Then
                    Dim bollingerHigh As Decimal = _bollingerHighPayload(signalCandle.PayloadDate)
                    Dim atr As Decimal = _atrPayload(signalCandle.PayloadDate)
                    Dim slPoint As Decimal = ConvertFloorCeling((bollingerHigh - entryPrice) / 2, _parentStrategy.TickSize, RoundOfType.Floor)
                    If slPoint < atr * _userInputs.ATRMultiplier Then
                        slPoint = ConvertFloorCeling(bollingerHigh - entryPrice, _parentStrategy.TickSize, RoundOfType.Floor)
                    End If
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                                            .Quantity = quantity,
                                                            .Stoploss = .EntryPrice - slPoint,
                                                            .Target = bollingerHigh,
                                                            .Buffer = 0,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.Market,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                            .Supporting2 = slPoint,
                                                            .Supporting3 = Math.Round(atr, 2)
                                                        }

                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                ElseIf signal.Item4 = Trade.TradeExecutionDirection.Sell Then
                    Dim bollingerLow As Decimal = _bollingerLowPayload(signalCandle.PayloadDate)
                    Dim atr As Decimal = _atrPayload(signalCandle.PayloadDate)
                    Dim slPoint As Decimal = ConvertFloorCeling((entryPrice - bollingerLow) / 2, _parentStrategy.TickSize, RoundOfType.Floor)
                    If slPoint < atr * _userInputs.ATRMultiplier Then
                        slPoint = ConvertFloorCeling(entryPrice - bollingerLow, _parentStrategy.TickSize, RoundOfType.Floor)
                    End If
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, _userInputs.MaxLossPerTrade, Trade.TypeOfStock.Cash)

                    Dim parameter As PlaceOrderParameters = New PlaceOrderParameters With {
                                                            .EntryPrice = entryPrice,
                                                            .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                                            .Quantity = quantity,
                                                            .Stoploss = .EntryPrice + slPoint,
                                                            .Target = bollingerLow,
                                                            .Buffer = 0,
                                                            .SignalCandle = signalCandle,
                                                            .OrderType = Trade.TypeOfOrder.Market,
                                                            .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                                            .Supporting2 = slPoint,
                                                            .Supporting3 = Math.Round(atr, 2)
                                                        }

                    ret = New Tuple(Of Boolean, List(Of PlaceOrderParameters))(True, New List(Of PlaceOrderParameters) From {parameter})
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForExitOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, String))
        Dim ret As Tuple(Of Boolean, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        'If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
        '    Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        '    Dim triggerPrice As Decimal = Decimal.MinValue
        '    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
        '        If currentMinuteCandle.PreviousCandlePayload.Low > _smaPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
        '            Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
        '            If currentTrade.EntryPrice + breakevenPoint < currentTick.Open Then
        '                triggerPrice = currentTrade.EntryPrice + breakevenPoint
        '            End If
        '        End If
        '    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
        '        If currentMinuteCandle.PreviousCandlePayload.High < _smaPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
        '            Dim breakevenPoint As Decimal = _parentStrategy.GetBreakevenPoint(_tradingSymbol, currentTrade.EntryPrice, currentTrade.Quantity, currentTrade.EntryDirection, currentTrade.LotSize, currentTrade.StockType)
        '            If currentTrade.EntryPrice - breakevenPoint > currentTick.Open Then
        '                triggerPrice = currentTrade.EntryPrice - breakevenPoint
        '            End If
        '        End If
        '    End If
        '    If triggerPrice <> Decimal.MinValue AndAlso triggerPrice <> currentTrade.PotentialStopLoss Then
        '        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, String.Format("Move to breakeven at {0}", currentTick.PayloadDate.ToString("HH:mm:ss")))
        '    End If
        'End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyTargetOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            Dim price As Decimal = Decimal.MinValue
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                Dim bollingerHigh As Decimal = _bollingerHighPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                If bollingerHigh - currentMinuteCandle.PreviousCandlePayload.High > 0 Then
                    price = ConvertFloorCeling(bollingerHigh + (bollingerHigh - currentMinuteCandle.PreviousCandlePayload.High), _parentStrategy.TickSize, RoundOfType.Celing)
                Else
                    price = ConvertFloorCeling(currentMinuteCandle.PreviousCandlePayload.High, _parentStrategy.TickSize, RoundOfType.Celing)
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                Dim bollingerLow As Decimal = _bollingerLowPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate)
                If currentMinuteCandle.PreviousCandlePayload.Low - bollingerLow > 0 Then
                    price = ConvertFloorCeling(bollingerLow - (currentMinuteCandle.PreviousCandlePayload.Low - bollingerLow), _parentStrategy.TickSize, RoundOfType.Floor)
                Else
                    price = ConvertFloorCeling(currentMinuteCandle.PreviousCandlePayload.Low, _parentStrategy.TickSize, RoundOfType.Floor)
                End If
            End If
            If price <> Decimal.MinValue AndAlso price <> currentTrade.PotentialTarget Then
                ret = New Tuple(Of Boolean, Decimal, String)(True, price, Math.Abs(price - currentTrade.EntryPrice))
            End If
        End If
        Return ret
    End Function

    Public Overrides Function IsTriggerReceivedForExitCNCEODOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Throw New NotImplementedException()
    End Function

    Public Overrides Async Function UpdateRequiredCollectionsAsync(currentTick As Payload) As Task
        Await Task.Delay(0).ConfigureAwait(False)
    End Function

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            If _direction = Trade.TradeExecutionDirection.Buy Then
                If currentCandle.PreviousCandlePayload.Close < _bollingerLowPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, currentTick.Open, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Buy)
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                If currentCandle.PreviousCandlePayload.Close > _bollingerHighPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, Decimal, Payload, Trade.TradeExecutionDirection)(True, currentTick.Open, currentCandle.PreviousCandlePayload, Trade.TradeExecutionDirection.Sell)
                End If
            End If
        End If
        Return ret
    End Function
End Class
Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class PreviousDayHKTrendSwingStrategy
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
        Public ATRMultiplier As Decimal
        Public TargetMultiplier As Decimal
    End Class
#End Region

    Private _swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing

    Private _dayATR As Decimal = Decimal.MinValue
    Private _direction As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None

    Private ReadOnly _userInputs As StrategyRuleEntities
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

        Dim eodPayload As Dictionary(Of Date, Payload) = _parentStrategy.Cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Cash, _tradingSymbol, _tradingDate.AddYears(-1), _tradingDate)
        If eodPayload IsNot Nothing AndAlso eodPayload.Count > 100 AndAlso eodPayload.ContainsKey(_tradingDate.Date) Then
            Dim atrPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.ATR.CalculateATR(14, eodPayload, atrPayload)
            _dayATR = atrPayload(_tradingDate.Date)

            Dim hkPayload As Dictionary(Of Date, Payload) = Nothing
            Indicator.HeikenAshi.ConvertToHeikenAshi(eodPayload, hkPayload)
            If hkPayload(_tradingDate.Date).CandleColor = Color.Green Then
                _direction = Trade.TradeExecutionDirection.Buy
            ElseIf hkPayload(_tradingDate.Date).CandleColor = Color.Red Then
                _direction = Trade.TradeExecutionDirection.Sell
            End If
        End If

        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastTrade As Trade = _parentStrategy.GetLastTradeOfTheStock(currentMinuteCandle, Trade.TypeOfTrade.MIS)
                If lastTrade IsNot Nothing Then
                    If lastTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate Then
                        signalCandle = signal.Item4
                    End If
                Else
                    signalCandle = signal.Item4
                End If
            End If

            If signalCandle IsNot Nothing Then
                Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signal.Item2, RoundOfType.Floor)
                If signal.Item5 = Trade.TradeExecutionDirection.Buy Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signal.Item3
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, stoploss, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
                            }
                ElseIf signal.Item5 = Trade.TradeExecutionDirection.Sell Then
                    Dim entryPrice As Decimal = signal.Item2
                    Dim stoploss As Decimal = signal.Item3
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, stoploss, entryPrice, Math.Abs(_userInputs.MaxLossPerTrade) * -1, Trade.TypeOfStock.Cash)
                    Dim target As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, Math.Abs(_userInputs.MaxLossPerTrade) * _userInputs.TargetMultiplier, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)

                    parameter = New PlaceOrderParameters With {
                                .EntryPrice = entryPrice,
                                .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                .Quantity = quantity,
                                .Stoploss = stoploss,
                                .Target = target,
                                .Buffer = buffer,
                                .SignalCandle = signalCandle,
                                .OrderType = Trade.TypeOfOrder.SL,
                                .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss")
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
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                If currentMinuteCandle.PreviousCandlePayload.Low <= currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                If currentMinuteCandle.PreviousCandlePayload.High >= currentTrade.PotentialStopLoss Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
            If ret Is Nothing Then
                Dim signal As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = GetEntrySignal(currentMinuteCandle, currentTick)
                If signal IsNot Nothing AndAlso signal.Item1 Then
                    If currentTrade.EntryDirection <> signal.Item5 OrElse
                        currentTrade.EntryPrice <> signal.Item2 OrElse
                        currentTrade.SignalCandle.PayloadDate <> signal.Item4.PayloadDate Then
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

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)
        Dim ret As Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim candle As Payload = currentCandle.PreviousCandlePayload
            Dim swing As Indicator.Swing = _swingPayload(candle.PayloadDate)
            If _direction = Trade.TradeExecutionDirection.Buy Then
                Dim signalCandle As Payload = _signalPayload(swing.SwingHighTime)
                If signalCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(swing.SwingHigh, RoundOfType.Floor)
                    Dim entry As Decimal = swing.SwingHigh + buffer
                    Dim stoploss As Decimal = Math.Min(signalCandle.Low, signalCandle.PreviousCandlePayload.Low)
                    Dim nextCandle As Payload = _signalPayload(swing.SwingHighTime.AddMinutes(_parentStrategy.SignalTimeFrame))
                    If nextCandle.Low < stoploss Then stoploss = nextCandle.Low
                    stoploss = stoploss - buffer
                    If entry - stoploss < _dayATR * _userInputs.ATRMultiplier Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)(True, entry, stoploss, signalCandle, Trade.TradeExecutionDirection.Buy)
                    End If
                End If
            ElseIf _direction = Trade.TradeExecutionDirection.Sell Then
                Dim signalCandle As Payload = _signalPayload(swing.SwingLowTime)
                If signalCandle.PreviousCandlePayload.PayloadDate.Date = _tradingDate.Date Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(swing.SwingLow, RoundOfType.Floor)
                    Dim entry As Decimal = swing.SwingLow - buffer
                    Dim stoploss As Decimal = Math.Max(signalCandle.High, signalCandle.PreviousCandlePayload.High)
                    Dim nextCandle As Payload = _signalPayload(swing.SwingLowTime.AddMinutes(_parentStrategy.SignalTimeFrame))
                    If nextCandle.High > stoploss Then stoploss = nextCandle.High
                    stoploss = stoploss + buffer
                    If stoploss - entry < _dayATR * _userInputs.ATRMultiplier Then
                        ret = New Tuple(Of Boolean, Decimal, Decimal, Payload, Trade.TradeExecutionDirection)(True, entry, stoploss, signalCandle, Trade.TradeExecutionDirection.Sell)
                    End If
                End If
            End If
        End If
        Return ret
    End Function
End Class
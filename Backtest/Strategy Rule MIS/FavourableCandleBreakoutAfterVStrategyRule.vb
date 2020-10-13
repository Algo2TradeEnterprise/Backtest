Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class FavourableCandleBreakoutAfterVStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MaxLossPerTrade As Decimal
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _smaPayload As Dictionary(Of Date, Decimal) = Nothing
    Private _swingPayload As Dictionary(Of Date, Indicator.Swing) = Nothing

    Private _previousTradingDay As Date

    Private _slPoint As Decimal

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

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
        Indicator.SMA.CalculateSMA(50, Payload.PayloadFields.Close, _signalPayload, _smaPayload)
        Indicator.SwingHighLow.CalculateSwingHighLow(_signalPayload, False, _swingPayload)
        _previousTradingDay = _parentStrategy.Cmn.GetPreviousTradingDay(Common.DataBaseTable.Intraday_Cash, _tradingSymbol, _tradingDate)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandle IsNot Nothing AndAlso currentMinuteCandle.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            currentMinuteCandle.PayloadDate >= _tradeStartTime AndAlso Me.EligibleToTakeTrade AndAlso
            Not _parentStrategy.IsAnyTradeOfTheStockTargetReached(currentMinuteCandle, Trade.TypeOfTrade.MIS) Then
            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date) = GetEntrySignal(currentMinuteCandle, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                Dim lastExecutedOrder As Trade = _parentStrategy.GetLastExitTradeOfTheStock(currentMinuteCandle, _parentStrategy.TradeType)
                If lastExecutedOrder IsNot Nothing Then
                    Dim exitCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(lastExecutedOrder.ExitTime, _signalPayload))
                    If signal.Item4 >= exitCandle.PreviousCandlePayload.PayloadDate Then
                        If lastExecutedOrder.EntryDirection = signal.Item3 Then
                            Dim lastOrderSwingTime As Date = Date.ParseExact(lastExecutedOrder.Supporting2, "dd-MM-yyyy HH:mm:ss", Nothing)
                            If lastOrderSwingTime <> signal.Item4 Then
                                signalCandle = signal.Item2
                            End If
                        Else
                            signalCandle = signal.Item2
                        End If
                    End If
                Else
                    signalCandle = signal.Item2
                    _slPoint = ConvertFloorCeling(GetAverageHighestATR(signalCandle), _parentStrategy.TickSize, RoundOfType.Celing)
                End If
            End If

            If signalCandle IsNot Nothing Then
                If signal.Item3 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.High + buffer
                    Dim slPoint As Decimal = _slPoint
                    Dim lossPL As Decimal = _userInputs.MaxLossPerTrade + _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice, entryPrice - slPoint, lossPL, Trade.TypeOfStock.Cash)
                    Dim profitPL As Decimal = Math.Abs(_userInputs.MaxLossPerTrade) - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    Dim targetPoint As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, Trade.TradeExecutionDirection.Buy, Trade.TypeOfStock.Cash) - entryPrice

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
                                    .Supporting2 = signal.Item4.ToString("dd-MM-yyyy HH:mm:ss"),
                                    .Supporting3 = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
                                }
                ElseIf signal.Item3 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    Dim entryPrice As Decimal = signalCandle.Low - buffer
                    Dim slPoint As Decimal = _slPoint
                    Dim lossPL As Decimal = _userInputs.MaxLossPerTrade + _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    Dim quantity As Integer = _parentStrategy.CalculateQuantityFromTargetSL(_tradingSymbol, entryPrice + slPoint, entryPrice, lossPL, Trade.TypeOfStock.Cash)
                    Dim profitPL As Decimal = Math.Abs(_userInputs.MaxLossPerTrade) - _parentStrategy.StockPLAfterBrokerage(_tradingDate, _tradingSymbol)
                    Dim targetPoint As Decimal = entryPrice - _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, entryPrice, quantity, profitPL, Trade.TradeExecutionDirection.Sell, Trade.TypeOfStock.Cash)

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
                                    .Supporting2 = signal.Item4.ToString("dd-MM-yyyy HH:mm:ss"),
                                    .Supporting3 = _parentStrategy.StockNumberOfTrades(_tradingDate, _tradingSymbol) + 1
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
        If currentTrade IsNot Nothing Then
            Dim currentMinuteCandle As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
            If currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
                Dim direction As Trade.TradeExecutionDirection = GetDirection(currentMinuteCandle)
                If direction <> Trade.TradeExecutionDirection.None AndAlso direction <> currentTrade.EntryDirection Then
                    ret = New Tuple(Of Boolean, String)(True, "Opposite Direction Signal")
                End If
            ElseIf currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Open Then
                If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                    currentMinuteCandle.PreviousCandlePayload.Close < _smaPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                    currentMinuteCandle.PreviousCandlePayload.Close > _smaPayload(currentMinuteCandle.PreviousCandlePayload.PayloadDate) Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
                If ret Is Nothing Then
                    Dim signal As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date) = GetEntrySignal(currentMinuteCandle, currentTick)
                    If signal IsNot Nothing AndAlso signal.Item1 Then
                        If currentTrade.SignalCandle.PayloadDate <> signal.Item2.PayloadDate Then
                            ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                        End If
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

    Private Function GetDirection(ByVal currentCandle As Payload) As Trade.TradeExecutionDirection
        Dim ret As Trade.TradeExecutionDirection = Trade.TradeExecutionDirection.None
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            If currentCandle.PreviousCandlePayload.Close > _smaPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                Dim swingDetails As Indicator.Swing = _swingPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                If swingDetails.SwingLow > _smaPayload(swingDetails.SwingLowTime) AndAlso swingDetails.SwingLowTime.Date = _tradingDate.Date Then
                    ret = Trade.TradeExecutionDirection.Buy
                End If
            ElseIf currentCandle.PreviousCandlePayload.Close < _smaPayload(currentCandle.PreviousCandlePayload.PayloadDate) Then
                Dim swingDetails As Indicator.Swing = _swingPayload(currentCandle.PreviousCandlePayload.PayloadDate)
                If swingDetails.SwingHigh < _smaPayload(swingDetails.SwingHighTime) AndAlso swingDetails.SwingHighTime.Date = _tradingDate.Date Then
                    ret = Trade.TradeExecutionDirection.Sell
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetEntrySignal(ByVal currentCandle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)
        Dim ret As Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date) = Nothing
        If currentCandle IsNot Nothing AndAlso currentCandle.PreviousCandlePayload IsNot Nothing Then
            Dim signalCandle As Payload = currentCandle
            Dim direction As Trade.TradeExecutionDirection = GetDirection(currentCandle)
            Dim preDirection As Trade.TradeExecutionDirection = GetDirection(currentCandle.PreviousCandlePayload)
            If preDirection = direction Then signalCandle = currentCandle.PreviousCandlePayload
            If direction = Trade.TradeExecutionDirection.Buy Then
                If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate > _swingPayload(signalCandle.PreviousCandlePayload.PayloadDate).SwingLowTime Then
                    If currentCandle.PreviousCandlePayload.High < currentCandle.PreviousCandlePayload.PreviousCandlePayload.High Then
                        ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)(True, currentCandle.PreviousCandlePayload, direction, _swingPayload(signalCandle.PreviousCandlePayload.PayloadDate).SwingLowTime)
                    End If
                End If
            ElseIf direction = Trade.TradeExecutionDirection.Sell Then
                If currentCandle.PreviousCandlePayload.PreviousCandlePayload.PayloadDate > _swingPayload(signalCandle.PreviousCandlePayload.PayloadDate).SwingHighTime Then
                    If currentCandle.PreviousCandlePayload.Low > currentCandle.PreviousCandlePayload.PreviousCandlePayload.Low Then
                        ret = New Tuple(Of Boolean, Payload, Trade.TradeExecutionDirection, Date)(True, currentCandle.PreviousCandlePayload, direction, _swingPayload(signalCandle.PreviousCandlePayload.PayloadDate).SwingHighTime)
                    End If
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetAverageHighestATR(ByVal signalCandle As Payload) As Decimal
        Dim ret As Decimal = Decimal.MinValue
        If _atrPayload IsNot Nothing AndAlso _atrPayload.Count > 0 AndAlso signalCandle IsNot Nothing Then
            Dim todayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                 If x.Key.Date = _tradingDate.Date AndAlso x.Key <= signalCandle.PayloadDate Then
                                                                     Return x.Value
                                                                 Else
                                                                     Return Decimal.MinValue
                                                                 End If
                                                             End Function)

            Dim previousDayHighestATR As Decimal = _atrPayload.Max(Function(x)
                                                                       If x.Key.Date = _previousTradingDay.Date Then
                                                                           Return x.Value
                                                                       Else
                                                                           Return Decimal.MinValue
                                                                       End If
                                                                   End Function)

            If todayHighestATR <> Decimal.MinValue AndAlso previousDayHighestATR <> Decimal.MinValue Then
                ret = (todayHighestATR + previousDayHighestATR) / 2
            End If
        End If
        Return ret
    End Function
End Class
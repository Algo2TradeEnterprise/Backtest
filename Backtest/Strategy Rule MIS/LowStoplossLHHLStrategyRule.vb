Imports Backtest.StrategyHelper
Imports System.Threading
Imports Algo2TradeBLL
Imports Utilities.Numbers.NumberManipulation

Public Class LowStoplossLHHLStrategyRule
    Inherits StrategyRule

#Region "Entity"
    Enum StoplossMakeupType
        SingleLossMakeup = 1
        AllLossMakeup
        None
    End Enum

    Public Class StrategyRuleEntities
        Inherits RuleEntities

        Public MinimumInvestmentPerStock As Decimal
        Public MinStoplossPerTrade As Decimal
        Public MaxStoplossPerTrade As Decimal
        Public MaxProfitPerTrade As Decimal
        Public TypeOfSLMakeup As StoplossMakeupType
    End Class
#End Region

    Private _atrPayload As Dictionary(Of Date, Decimal)
    Private _userInputs As StrategyRuleEntities
    Private _quantity As Integer

    Public Sub New(ByVal inputPayload As Dictionary(Of Date, Payload),
                   ByVal lotSize As Integer,
                   ByVal parentStrategy As Strategy,
                   ByVal tradingDate As Date,
                   ByVal tradingSymbol As String,
                   ByVal canceller As CancellationTokenSource,
                   ByVal entities As RuleEntities)
        MyBase.New(inputPayload, lotSize, parentStrategy, tradingDate, tradingSymbol, canceller, entities)
        _userInputs = entities
        _quantity = Integer.MinValue
    End Sub

    Public Overrides Sub CompletePreProcessing()
        MyBase.CompletePreProcessing()

        Indicator.ATR.CalculateATR(14, _signalPayload, _atrPayload)
    End Sub

    Public Overrides Async Function IsTriggerReceivedForPlaceOrderAsync(currentTick As Payload) As Task(Of Tuple(Of Boolean, List(Of PlaceOrderParameters)))
        Dim ret As Tuple(Of Boolean, List(Of PlaceOrderParameters)) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        Dim currentMinuteCandlePayload As Payload = _signalPayload(_parentStrategy.GetCurrentXMinuteCandleTime(currentTick.PayloadDate, _signalPayload))
        Dim tradeStartTime As Date = New Date(_tradingDate.Year, _tradingDate.Month, _tradingDate.Day, _parentStrategy.TradeStartTime.Hours, _parentStrategy.TradeStartTime.Minutes, _parentStrategy.TradeStartTime.Seconds)
        Dim parameter As PlaceOrderParameters = Nothing
        If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.PreviousCandlePayload IsNot Nothing AndAlso
            Not _parentStrategy.IsTradeActive(currentTick, Trade.TypeOfTrade.MIS) AndAlso Not _parentStrategy.IsTradeOpen(currentTick, Trade.TypeOfTrade.MIS) AndAlso
            _parentStrategy.StockNumberOfTrades(currentTick.PayloadDate, currentTick.TradingSymbol) < Me._parentStrategy.NumberOfTradesPerStockPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) < _parentStrategy.OverAllProfitPerDay AndAlso
            _parentStrategy.TotalPLAfterBrokerage(currentTick.PayloadDate) > _parentStrategy.OverAllLossPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < _parentStrategy.StockMaxProfitPerDay AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(_parentStrategy.StockMaxLossPerDay) * -1 AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < Me.MaxProfitOfThisStock AndAlso
            _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) > Math.Abs(Me.MaxLossOfThisStock) * -1 AndAlso
            currentMinuteCandlePayload.PayloadDate >= tradeStartTime AndAlso Me.EligibleToTakeTrade Then

            If _quantity = Integer.MinValue Then
                _quantity = _parentStrategy.CalculateQuantityFromInvestment(LotSize, _userInputs.MinimumInvestmentPerStock, currentMinuteCandlePayload.PreviousCandlePayload.Close, _parentStrategy.StockType, True)
            End If

            Dim signalCandle As Payload = Nothing
            Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            Dim lastExecutedTrade As Trade = _parentStrategy.GetLastExecutedTradeOfTheStock(currentMinuteCandlePayload, _parentStrategy.TradeType)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If lastExecutedTrade Is Nothing Then
                    signalCandle = signal.Item3
                ElseIf lastExecutedTrade.SignalCandle.PayloadDate <> signal.Item3.PayloadDate Then
                    signalCandle = signal.Item3
                End If
            End If

            If signalCandle IsNot Nothing AndAlso signalCandle.PayloadDate < currentMinuteCandlePayload.PayloadDate Then
                Dim originalTargetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signalCandle.Open, _quantity, Math.Abs(_userInputs.MaxProfitPerTrade), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                Dim targetPoint As Decimal = originalTargetPrice - signalCandle.Open
                Dim targetRemark As String = "Original Target"
                If _userInputs.TypeOfSLMakeup = StoplossMakeupType.SingleLossMakeup Then
                    If lastExecutedTrade IsNot Nothing AndAlso lastExecutedTrade.ExitCondition = Trade.TradeExitCondition.StopLoss Then
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signalCandle.Open, _quantity, Math.Abs(lastExecutedTrade.PLAfterBrokerage), Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        targetPoint = targetPrice - signalCandle.Open
                        targetRemark = "SL Makeup Trade"
                    End If
                ElseIf _userInputs.TypeOfSLMakeup = StoplossMakeupType.AllLossMakeup Then
                    If _parentStrategy.StockPLAfterBrokerage(currentTick.PayloadDate, currentTick.TradingSymbol) < 0 Then
                        Dim targetPrice As Decimal = _parentStrategy.CalculatorTargetOrStoploss(_tradingSymbol, signalCandle.Open, _quantity, (_userInputs.MinStoplossPerTrade + _userInputs.MaxStoplossPerTrade) / 2, Trade.TradeExecutionDirection.Buy, _parentStrategy.StockType)
                        targetPoint = targetPrice - signalCandle.Open
                        targetRemark = "SL Makeup Trade"
                    End If
                End If

                If signal.Item2 = Trade.TradeExecutionDirection.Buy Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.High, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = signalCandle.High + buffer,
                                    .EntryDirection = Trade.TradeExecutionDirection.Buy,
                                    .Quantity = _quantity,
                                    .Stoploss = .EntryPrice - signal.Item4,
                                    .Target = .EntryPrice + targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = targetRemark,
                                    .Supporting3 = targetPoint,
                                    .Supporting4 = signal.Item4
                                }
                ElseIf signal.Item2 = Trade.TradeExecutionDirection.Sell Then
                    Dim buffer As Decimal = _parentStrategy.CalculateBuffer(signalCandle.Low, RoundOfType.Floor)
                    parameter = New PlaceOrderParameters With {
                                    .EntryPrice = signalCandle.Low,
                                    .EntryDirection = Trade.TradeExecutionDirection.Sell,
                                    .Quantity = _quantity,
                                    .Stoploss = .EntryPrice + signal.Item4,
                                    .Target = .EntryPrice - targetPoint,
                                    .Buffer = buffer,
                                    .SignalCandle = signalCandle,
                                    .OrderType = Trade.TypeOfOrder.SL,
                                    .Supporting1 = signalCandle.PayloadDate.ToString("HH:mm:ss"),
                                    .Supporting2 = targetRemark,
                                    .Supporting3 = targetPoint,
                                    .Supporting4 = signal.Item4
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
            Dim signal As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal) = GetEntrySignal(currentMinuteCandlePayload.PreviousCandlePayload, currentTick)
            If signal IsNot Nothing AndAlso signal.Item1 Then
                If currentTrade.EntryDirection <> signal.Item2 OrElse signal.Item3.PayloadDate <> currentTrade.SignalCandle.PayloadDate Then
                    ret = New Tuple(Of Boolean, String)(True, "Invalid Signal")
                End If
            End If
        End If
        Return ret
    End Function

    Public Overrides Async Function IsTriggerReceivedForModifyStoplossOrderAsync(currentTick As Payload, currentTrade As Trade) As Task(Of Tuple(Of Boolean, Decimal, String))
        Dim ret As Tuple(Of Boolean, Decimal, String) = Nothing
        Await Task.Delay(0).ConfigureAwait(False)
        'If currentTrade IsNot Nothing AndAlso currentTrade.TradeCurrentStatus = Trade.TradeExecutionStatus.Inprogress Then
        '    Dim slPoint As Decimal = CDec(currentTrade.Supporting4)
        '    Dim triggerPrice As Decimal = Decimal.MinValue
        '    If currentTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
        '        triggerPrice = currentTrade.EntryPrice - slPoint
        '    ElseIf currentTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
        '        triggerPrice = currentTrade.EntryPrice + slPoint
        '    End If
        '    If triggerPrice <> Decimal.MinValue AndAlso currentTrade.PotentialStopLoss <> triggerPrice Then
        '        ret = New Tuple(Of Boolean, Decimal, String)(True, triggerPrice, slPoint)
        '    End If
        'End If
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

    Private Function GetEntrySignal(ByVal candle As Payload, ByVal currentTick As Payload) As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal)
        Dim ret As Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal) = Nothing
        If candle IsNot Nothing AndAlso candle.PreviousCandlePayload IsNot Nothing AndAlso
            Not candle.DeadCandle AndAlso Not candle.PreviousCandlePayload.DeadCandle Then
            Dim buySLPriceAndPoint As Tuple(Of Decimal, Decimal) = GetStoplossPriceAndPoint(candle, Trade.TradeExecutionDirection.Buy)
            Dim sellSLPriceAndPoint As Tuple(Of Decimal, Decimal) = GetStoplossPriceAndPoint(candle, Trade.TradeExecutionDirection.Sell)
            Dim atr As Decimal = ConvertFloorCeling(_atrPayload(candle.PayloadDate), _parentStrategy.TickSize, RoundOfType.Floor)
            If candle.High < candle.PreviousCandlePayload.High AndAlso candle.Low > candle.PreviousCandlePayload.Low Then
                If Math.Abs(buySLPriceAndPoint.Item1) >= Math.Abs(_userInputs.MinStoplossPerTrade) AndAlso
                    Math.Abs(buySLPriceAndPoint.Item1) <= Math.Abs(_userInputs.MaxStoplossPerTrade) AndAlso
                    buySLPriceAndPoint.Item2 <= atr AndAlso
                    Math.Abs(sellSLPriceAndPoint.Item1) >= Math.Abs(_userInputs.MinStoplossPerTrade) AndAlso
                    Math.Abs(sellSLPriceAndPoint.Item1) <= Math.Abs(_userInputs.MaxStoplossPerTrade) AndAlso
                    sellSLPriceAndPoint.Item2 <= atr Then
                    Dim middlePoint As Decimal = (candle.High + candle.Low) / 2
                    Dim range As Decimal = candle.High - middlePoint
                    If currentTick.Open > middlePoint + range * 50 / 100 Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal)(True, Trade.TradeExecutionDirection.Buy, candle, buySLPriceAndPoint.Item2)
                    ElseIf currentTick.Open < middlePoint - range * 50 / 100 Then
                        ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal)(True, Trade.TradeExecutionDirection.Sell, candle, sellSLPriceAndPoint.Item2)
                    End If
                ElseIf Math.Abs(buySLPriceAndPoint.Item1) >= Math.Abs(_userInputs.MinStoplossPerTrade) AndAlso
                    Math.Abs(buySLPriceAndPoint.Item1) <= Math.Abs(_userInputs.MaxStoplossPerTrade) AndAlso
                    buySLPriceAndPoint.Item2 <= atr Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal)(True, Trade.TradeExecutionDirection.Buy, candle, buySLPriceAndPoint.Item2)
                ElseIf Math.Abs(sellSLPriceAndPoint.Item1) >= Math.Abs(_userInputs.MinStoplossPerTrade) AndAlso
                    Math.Abs(sellSLPriceAndPoint.Item1) <= Math.Abs(_userInputs.MaxStoplossPerTrade) AndAlso
                    sellSLPriceAndPoint.Item2 <= atr Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal)(True, Trade.TradeExecutionDirection.Sell, candle, sellSLPriceAndPoint.Item2)
                Else
                    Console.WriteLine(String.Format("Trade neglected because of SL price. Signal candle,{0}, Buy SL Amount:{1}, Sell SL Amount:{2}, Trading Symbol:{3}",
                                                    candle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), buySLPriceAndPoint.Item1, sellSLPriceAndPoint.Item1, candle.TradingSymbol))
                End If
            ElseIf candle.High < candle.PreviousCandlePayload.High Then
                If Math.Abs(buySLPriceAndPoint.Item1) >= Math.Abs(_userInputs.MinStoplossPerTrade) AndAlso
                    Math.Abs(buySLPriceAndPoint.Item1) <= Math.Abs(_userInputs.MaxStoplossPerTrade) AndAlso
                    buySLPriceAndPoint.Item2 <= atr Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal)(True, Trade.TradeExecutionDirection.Buy, candle, buySLPriceAndPoint.Item2)
                Else
                    Console.WriteLine(String.Format("Trade neglected because of SL price. Signal candle,{0}, Buy SL Amount:{1}, Sell SL Amount:{2}, Trading Symbol:{3}",
                                                    candle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), buySLPriceAndPoint.Item1, sellSLPriceAndPoint.Item1, candle.TradingSymbol))
                End If
            ElseIf candle.Low > candle.PreviousCandlePayload.Low Then
                If Math.Abs(sellSLPriceAndPoint.Item1) >= Math.Abs(_userInputs.MinStoplossPerTrade) AndAlso
                    Math.Abs(sellSLPriceAndPoint.Item1) <= Math.Abs(_userInputs.MaxStoplossPerTrade) AndAlso
                    sellSLPriceAndPoint.Item2 <= atr Then
                    ret = New Tuple(Of Boolean, Trade.TradeExecutionDirection, Payload, Decimal)(True, Trade.TradeExecutionDirection.Sell, candle, sellSLPriceAndPoint.Item2)
                Else
                    Console.WriteLine(String.Format("Trade neglected because of SL price. Signal candle,{0}, Buy SL Amount:{1}, Sell SL Amount:{2}, Trading Symbol:{3}",
                                                    candle.PayloadDate.ToString("dd-MM-yyyy HH:mm:ss"), buySLPriceAndPoint.Item1, sellSLPriceAndPoint.Item1, candle.TradingSymbol))
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetStoplossPriceAndPoint(ByVal candle As Payload, ByVal direction As Trade.TradeExecutionDirection) As Tuple(Of Decimal, Decimal)
        Dim ret As Tuple(Of Decimal, Decimal) = Nothing
        If direction = Trade.TradeExecutionDirection.Buy Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.High, RoundOfType.Floor)
            Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, candle.High + buffer, candle.Low - buffer, _quantity, LotSize, _parentStrategy.StockType)
            ret = New Tuple(Of Decimal, Decimal)(pl, candle.High - candle.Low + 2 * buffer)
        ElseIf direction = Trade.TradeExecutionDirection.Sell Then
            Dim buffer As Decimal = _parentStrategy.CalculateBuffer(candle.Low, RoundOfType.Floor)
            Dim pl As Decimal = _parentStrategy.CalculatePL(_tradingSymbol, candle.High + buffer, candle.Low - buffer, _quantity, LotSize, _parentStrategy.StockType)
            ret = New Tuple(Of Decimal, Decimal)(pl, candle.High - candle.Low + 2 * buffer)
        End If
        Return ret
    End Function
End Class